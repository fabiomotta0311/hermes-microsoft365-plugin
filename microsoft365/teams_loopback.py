"""Teams loopback: turn an admitted Activity into a reply, without trusting its body for routing.

The gap this closes
-------------------
:mod:`microsoft365.teams_inbound` decides which session an incoming Activity belongs to. Nothing
yet turns that decision into an outgoing reply, so the Teams story is still half a conversation:
the plugin can be *read to*, not *answered*.

Why this is a separate module and not part of the boundary
---------------------------------------------------------
Admitting an Activity must never be able to send anything. If the boundary could reply, then a
webhook that is merely misconfigured -- not even forged -- would have an outbound path. Keeping
"may this reach a session?" and "may this session answer?" as two separately-constructed
decisions is what preserves that. This module only runs when a caller has deliberately built it.

The routing rule, which is the whole security story
--------------------------------------------------
**The destination chat never comes from the Activity.** An Activity body carries a
``conversation.id`` and a ``serviceUrl``, and both are attacker-controlled in the threat model
that matters: anyone who can reach the webhook can put anything there. Routing a reply by those
values would let a forged body redirect a session's answer to an attacker-chosen conversation.

So the destination comes from the operator's own binding -- `ReplyRoute`, registered alongside
the identity -- and the Activity is only ever *read for its text*. A missing route is a refusal,
never a fallback to the body.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Mapping

from .errors import MESSAGES, GraphError
from .teams_inbound import Activity, Identity, InboundDecision, InboundTeamsError

#: Replies are bounded like any other outbound text in this plugin.
MAX_REPLY_LENGTH = 4000

#: Teams' own limit on a chat message body. Refused here so a too-long reply is a clear refusal
#: rather than a truncation the caller never asked for.
TEAMS_MESSAGE_LIMIT = 28_000

NO_REPLY_ROUTE = "no_reply_route"
REPLY_NOT_DELIVERABLE = "reply_not_deliverable"
EMPTY_REPLY = "empty_reply"
REPLY_TOO_LONG = "reply_too_long"
REPLY_TEMPLATE_FAILED = "reply_template_failed"

_ACTIONS = {
    NO_REPLY_ROUTE: "bind a reply route for this identity before answering",
    REPLY_NOT_DELIVERABLE: "this activity is not something a reply can be sent to",
    EMPTY_REPLY: "compose a non-blank reply",
    REPLY_TOO_LONG: "shorten the reply to the teams message limit",
    REPLY_TEMPLATE_FAILED: "the reply template did not produce a usable plain-text body",
}


class LoopbackError(GraphError):
    """A sanitized loopback failure carrying an explicit, machine-readable ``status``."""

    def __init__(self, category: str, *, status: str, action: str | None = None):
        super().__init__(category, MESSAGES[category])
        self.status = str(status)
        self.action = action if action is not None else _ACTIONS.get(self.status)

    def to_result(self) -> dict:
        payload = super().to_result()
        payload["status"] = self.status
        if self.action:
            payload["action"] = self.action
        return payload


@dataclass(frozen=True)
class ReplyRoute:
    """Where a session's answer goes, as decided by the operator -- never by the Activity.

    ``chat_id`` is a Microsoft Graph chat id (what ``teams.send_chat_message`` takes). It is
    deliberately *not* the Bot Framework ``conversation.id``: they are different identifiers in
    different systems, and conflating them is the mistake this dataclass exists to prevent.
    """

    chat_id: str
    content_type: str = "text"

    def __post_init__(self) -> None:
        if not isinstance(self.chat_id, str) or not self.chat_id.strip():
            raise LoopbackError("validation_error", status=NO_REPLY_ROUTE)
        if self.content_type not in {"text", "html"}:
            raise LoopbackError("validation_error", status=NO_REPLY_ROUTE)


class ReplyRouter:
    """Binds external identities to reply routes, and composes the reply for one Activity."""

    def __init__(self) -> None:
        self._routes: dict[tuple[str, str, str], ReplyRoute] = {}

    def bind(self, identity: Identity, route: ReplyRoute) -> None:
        if not isinstance(route, ReplyRoute):
            raise LoopbackError("validation_error", status=NO_REPLY_ROUTE)
        self._routes[identity.as_key()] = route

    def unbind(self, identity: Identity) -> bool:
        return self._routes.pop(identity.as_key(), None) is not None

    def route_for(self, identity: Identity) -> ReplyRoute:
        """The route bound to an identity, or a refusal.

        There is no default route and no "reply to the same conversation the Activity came
        from": an unbound identity is refused, because the body is not a routing authority.
        """
        route = self._routes.get(identity.as_key())
        if route is None:
            raise LoopbackError("validation_error", status=NO_REPLY_ROUTE)
        return route


def validate_reply_text(text: Any) -> str:
    """A non-blank, bounded reply body, or a refusal."""
    if not isinstance(text, str):
        raise LoopbackError("validation_error", status=EMPTY_REPLY)
    stripped = text.strip()
    if not stripped:
        raise LoopbackError("validation_error", status=EMPTY_REPLY)
    if len(stripped) > TEAMS_MESSAGE_LIMIT:
        raise LoopbackError("validation_error", status=REPLY_TOO_LONG)
    return stripped


@dataclass(frozen=True)
class ReplyPlan:
    """A composed reply, ready to hand to the already-authorised send path."""

    chat_id: str
    body: str
    content_type: str
    session_id: str
    activity_id: str

    def to_send_arguments(self) -> dict:
        """The exact arguments ``teams.send_chat_message`` takes -- nothing more.

        Building the call's arguments here, rather than letting a caller assemble them, is what
        keeps the route authoritative: there is no parameter this function accepts that could
        redirect the send.
        """
        return {"chat_id": self.chat_id, "body": self.body, "content_type": self.content_type}

    def to_result(self) -> dict:
        """The rendered envelope, without the reply text or the destination chat id."""
        return {
            "deliverable": True,
            "content_type": self.content_type,
            "body_length": len(self.body),
            "activity_id": self.activity_id,
        }


def plan_reply(
    *,
    decision: InboundDecision,
    route: ReplyRoute,
    text: Any,
) -> ReplyPlan:
    """Compose the reply for one admitted Activity, refusing anything not deliverable.

    A ``message`` Activity is what a reply answers. A lifecycle Activity (``conversationUpdate``,
    ``typing``, ``messageReaction``) is admitted by the boundary so a session can *learn* about
    it, but it is not a turn: replying to it would send a message nobody asked for.
    """
    if not isinstance(decision, InboundDecision):
        raise LoopbackError("validation_error", status=REPLY_NOT_DELIVERABLE)
    if decision.duplicate:
        # A retried delivery was already answered. Replying again is how a retry storm becomes a
        # duplicate-message storm in the user's Teams client.
        raise LoopbackError("validation_error", status=REPLY_NOT_DELIVERABLE)
    if decision.activity.type != "message":
        raise LoopbackError("validation_error", status=REPLY_NOT_DELIVERABLE)
    if decision.activity.activity_id is None:
        raise LoopbackError("validation_error", status=REPLY_NOT_DELIVERABLE)

    body = validate_reply_text(text)
    if route.content_type == "text" and len(body) > MAX_REPLY_LENGTH:
        raise LoopbackError("validation_error", status=REPLY_TOO_LONG)

    return ReplyPlan(
        chat_id=route.chat_id,
        body=body,
        content_type=route.content_type,
        session_id=decision.session_id,
        activity_id=decision.activity.activity_id,
    )


def plan_reply_for(
    *,
    decision: InboundDecision,
    router: ReplyRouter,
    render: Callable[[Activity], Any],
) -> ReplyPlan:
    """Resolve the route from the binding, run the renderer, and compose the reply.

    ``render`` receives the projected :class:`Activity` -- never the raw body -- and returns the
    reply text. Its failure is a refusal, not an empty reply: a session that could not answer
    must not appear to have answered nothing.
    """
    route = router.route_for(decision.identity)
    try:
        text = render(decision.activity)
    except Exception as exc:  # noqa: BLE001 - a failed render is a refusal, never an empty send
        raise LoopbackError("internal_error", status=REPLY_TEMPLATE_FAILED) from exc
    return plan_reply(decision=decision, route=route, text=text)


@dataclass(frozen=True)
class ReplyOutcome:
    """What happened when the plan was handed to the send path."""

    plan: ReplyPlan
    delivered: bool
    message_id: str | None = None
    error_status: str | None = None

    def to_result(self) -> dict:
        payload: dict[str, Any] = {
            "delivered": self.delivered,
            "activity_id": self.plan.activity_id,
            "body_length": len(self.plan.body),
        }
        if self.message_id:
            payload["message_id"] = self.message_id
        if self.error_status:
            payload["error_status"] = self.error_status
        return payload


def deliver_reply(plan: ReplyPlan, *, send: Callable[[Mapping], Mapping]) -> ReplyOutcome:
    """Hand the plan to the send path and report what it said.

    ``send`` is the already-authorised outbound operation -- in production, the
    ``teams.send_chat_message`` handler with its own settings, preflight and permission checks.
    This function passes the arguments the plan built and no others, so the send path cannot be
    steered by anything upstream of it.

    A send failure is reported as an un-delivered outcome with a sanitized status rather than
    raised: the caller has an Activity it already committed to answering and needs to know the
    answer did not go out, which is a different situation from the request never being made.
    """
    if not isinstance(plan, ReplyPlan):
        raise LoopbackError("validation_error", status=REPLY_NOT_DELIVERABLE)
    try:
        response = send(plan.to_send_arguments())
    except InboundTeamsError as exc:
        return ReplyOutcome(plan=plan, delivered=False, error_status=exc.status)
    except GraphError as exc:
        return ReplyOutcome(
            plan=plan, delivered=False, error_status=getattr(exc, "status", None) or "send_failed"
        )
    except Exception:  # noqa: BLE001 - an unexpected send failure is still not a delivered reply
        return ReplyOutcome(plan=plan, delivered=False, error_status="send_failed")

    if not isinstance(response, Mapping):
        return ReplyOutcome(plan=plan, delivered=False, error_status="send_failed")
    if response.get("status") == "succeeded":
        result = response.get("result")
        message_id = result.get("id") if isinstance(result, Mapping) else None
        return ReplyOutcome(
            plan=plan,
            delivered=True,
            message_id=message_id if isinstance(message_id, str) else None,
        )
    error = response.get("error")
    return ReplyOutcome(plan=plan, delivered=False, error_status=str(error or "send_failed"))