"""Offline tests for the Teams loopback (``microsoft365.teams_loopback``).

The claim under test: the destination of a reply is decided by the operator's binding and can
never be influenced by the Activity body. Every adversarial case here tries to make the body
steer the send, and every one must fail.
"""
from __future__ import annotations

import pytest

from microsoft365.teams_inbound import Activity, Identity, InboundDecision, InboundTeamsError
from microsoft365.teams_loopback import (
    NO_REPLY_ROUTE,
    REPLY_NOT_DELIVERABLE,
    REPLY_TEMPLATE_FAILED,
    REPLY_TOO_LONG,
    TEAMS_MESSAGE_LIMIT,
    LoopbackError,
    ReplyPlan,
    ReplyRoute,
    ReplyRouter,
    deliver_reply,
    plan_reply,
    plan_reply_for,
    validate_reply_text,
)

TENANT = "tenant-1"
SENDER = "sender-1"
CONVERSATION = "19:bf-conversation@thread.v2"
CHAT_ID = "19:graph-chat-id@thread.v2"
IDENTITY = Identity(TENANT, SENDER, CONVERSATION)


def activity(*, activity_type="message", text="hello", activity_id="activity-1", conversation=CONVERSATION):
    return Activity(
        type=activity_type,
        activity_id=activity_id,
        timestamp="2026-10-02T12:00:00.000Z",
        text=text,
        channel_id="msteams",
        tenant_id=TENANT,
        sender_id=SENDER,
        sender_name="Someone",
        recipient_id="28:bot",
        conversation_id=conversation,
        conversation_type="personal",
        service_url="https://smba.trafficmanager.net/amer/",
        locale="pt-BR",
    )


def decision(*, duplicate=False, model=None):
    return InboundDecision(
        session_id="session-a",
        identity=IDENTITY,
        activity=model or activity(),
        duplicate=duplicate,
    )


# ------------------------------------------------------------------ ReplyRoute


def test_a_route_needs_a_real_chat_id():
    assert ReplyRoute(chat_id=CHAT_ID).chat_id == CHAT_ID
    for bad in (None, "", "   ", 42):
        with pytest.raises(LoopbackError) as caught:
            ReplyRoute(chat_id=bad)
        assert caught.value.status == NO_REPLY_ROUTE


def test_a_route_rejects_an_unknown_content_type():
    with pytest.raises(LoopbackError) as caught:
        ReplyRoute(chat_id=CHAT_ID, content_type="markdown")
    assert caught.value.status == NO_REPLY_ROUTE


# ------------------------------------------------------------------ router


def test_an_unbound_identity_has_no_route():
    router = ReplyRouter()
    with pytest.raises(LoopbackError) as caught:
        router.route_for(IDENTITY)
    assert caught.value.status == NO_REPLY_ROUTE


def test_a_bound_identity_resolves_to_the_operators_route():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    assert router.route_for(IDENTITY).chat_id == CHAT_ID
    assert router.unbind(IDENTITY) is True
    assert router.unbind(IDENTITY) is False


def test_the_route_is_keyed_by_tenant_too():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    with pytest.raises(LoopbackError) as caught:
        router.route_for(Identity("tenant-2", SENDER, CONVERSATION))
    assert caught.value.status == NO_REPLY_ROUTE


# ------------------------------------------------------------------ the routing invariant


def test_the_reply_goes_to_the_bound_chat_and_never_to_the_activity_conversation():
    """The core guarantee: the body cannot steer the destination."""
    hostile = activity(conversation="19:attacker-controlled@thread.v2")
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))

    plan = plan_reply_for(decision=decision(model=hostile), router=router, render=lambda a: "ok")

    assert plan.chat_id == CHAT_ID
    assert plan.chat_id != hostile.conversation_id
    assert plan.to_send_arguments() == {"chat_id": CHAT_ID, "body": "ok", "content_type": "text"}


def test_the_send_arguments_contain_only_chat_id_body_and_content_type():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(decision=decision(), router=router, render=lambda a: "ok")
    assert set(plan.to_send_arguments()) == {"chat_id", "body", "content_type"}


def test_a_service_url_in_the_body_never_reaches_the_plan():
    hostile = activity()
    hostile = Activity(**{**hostile.__dict__, "service_url": "https://evil.example/hook"})
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(decision=decision(model=hostile), router=router, render=lambda a: "ok")
    assert "evil.example" not in str(plan.to_send_arguments())
    assert "evil.example" not in str(plan.to_result())


# ------------------------------------------------------------------ deliverability


def test_a_duplicated_activity_is_not_answered_again():
    """A retry was already answered; answering again duplicates the message in Teams."""
    with pytest.raises(LoopbackError) as caught:
        plan_reply(decision=decision(duplicate=True), route=ReplyRoute(chat_id=CHAT_ID), text="ok")
    assert caught.value.status == REPLY_NOT_DELIVERABLE


@pytest.mark.parametrize("kind", ["conversationUpdate", "typing", "messageReaction"])
def test_a_lifecycle_activity_is_not_answered(kind):
    """Admitted so a session can learn about it, but it is not a turn to reply to."""
    with pytest.raises(LoopbackError) as caught:
        plan_reply(
            decision=decision(model=activity(activity_type=kind)),
            route=ReplyRoute(chat_id=CHAT_ID),
            text="ok",
        )
    assert caught.value.status == REPLY_NOT_DELIVERABLE


def test_an_activity_without_an_id_is_not_answered():
    with pytest.raises(LoopbackError) as caught:
        plan_reply(
            decision=decision(model=activity(activity_id=None)),
            route=ReplyRoute(chat_id=CHAT_ID),
            text="ok",
        )
    assert caught.value.status == REPLY_NOT_DELIVERABLE


def test_a_non_decision_is_refused():
    with pytest.raises(LoopbackError) as caught:
        plan_reply(decision={"session_id": "x"}, route=ReplyRoute(chat_id=CHAT_ID), text="ok")
    assert caught.value.status == REPLY_NOT_DELIVERABLE


# ------------------------------------------------------------------ reply text


def test_a_blank_reply_is_refused():
    for bad in (None, "", "   ", "\n\t ", 42):
        with pytest.raises(LoopbackError) as caught:
            validate_reply_text(bad)
        assert caught.value.status == "empty_reply"


def test_a_reply_over_the_teams_limit_is_refused_not_truncated():
    with pytest.raises(LoopbackError) as caught:
        validate_reply_text("x" * (TEAMS_MESSAGE_LIMIT + 1))
    assert caught.value.status == REPLY_TOO_LONG


def test_a_reply_within_the_limit_is_kept_verbatim():
    text = "a reply with spaces  and punctuation!"
    assert validate_reply_text(text) == text


def test_a_renderer_failure_is_a_refusal_not_an_empty_message():
    """A session that could not answer must not look like it answered nothing."""
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))

    def explode(_activity):
        raise RuntimeError("template blew up: secret=abc")

    with pytest.raises(LoopbackError) as caught:
        plan_reply_for(decision=decision(), router=router, render=explode)
    assert caught.value.status == REPLY_TEMPLATE_FAILED
    assert "secret" not in str(caught.value.to_result())


def test_the_renderer_sees_the_projected_activity_only():
    seen = {}

    def render(model):
        seen["model"] = model
        return "ok"

    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan_reply_for(decision=decision(), router=router, render=render)
    assert isinstance(seen["model"], Activity)
    assert seen["model"].text == "hello"


# ------------------------------------------------------------------ delivery


def ok_send(**kwargs):
    def send(arguments):
        return {"operation": "teams.send_chat_message", "status": "succeeded",
                "result": {"id": "message-1", **arguments}}

    return send


def test_a_successful_send_reports_the_message_id():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(decision=decision(), router=router, render=lambda a: "ok")

    outcome = deliver_reply(plan, send=lambda arguments: {
        "status": "succeeded", "result": {"id": "message-1"}
    })

    assert outcome.delivered is True
    assert outcome.message_id == "message-1"
    assert outcome.to_result()["delivered"] is True


def test_the_send_receives_exactly_the_planned_arguments():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(decision=decision(), router=router, render=lambda a: "the answer")
    captured = {}

    def send(arguments):
        captured.update(arguments)
        return {"status": "succeeded", "result": {"id": "m-1"}}

    deliver_reply(plan, send=send)
    assert captured == {"chat_id": CHAT_ID, "body": "the answer", "content_type": "text"}


def test_a_graph_refusal_is_an_undelivered_outcome_not_an_exception():
    plan = ReplyPlan(chat_id=CHAT_ID, body="ok", content_type="text",
                     session_id="session-a", activity_id="activity-1")

    def send(arguments):
        return {"operation": "teams.send_chat_message", "status": "refused",
                "error": "permission_denied"}

    outcome = deliver_reply(plan, send=send)
    assert outcome.delivered is False
    assert outcome.error_status == "permission_denied"


def test_a_raising_send_is_an_undelivered_outcome():
    plan = ReplyPlan(chat_id=CHAT_ID, body="ok", content_type="text",
                     session_id="session-a", activity_id="activity-1")

    def send(arguments):
        raise InboundTeamsError("authentication_required", status="invalid_token")

    outcome = deliver_reply(plan, send=send)
    assert outcome.delivered is False
    assert outcome.error_status == "invalid_token"


def test_an_unexpected_send_failure_is_undelivered_and_sanitized():
    plan = ReplyPlan(chat_id=CHAT_ID, body="ok", content_type="text",
                     session_id="session-a", activity_id="activity-1")

    def send(arguments):
        raise RuntimeError("boom: token=abc123 at https://graph.microsoft.com")

    outcome = deliver_reply(plan, send=send)
    assert outcome.delivered is False
    assert outcome.error_status == "send_failed"
    assert "abc123" not in str(outcome.to_result())
    assert "graph.microsoft.com" not in str(outcome.to_result())


def test_a_non_mapping_send_response_is_undelivered():
    plan = ReplyPlan(chat_id=CHAT_ID, body="ok", content_type="text",
                     session_id="session-a", activity_id="activity-1")
    outcome = deliver_reply(plan, send=lambda arguments: "sure, sent it")
    assert outcome.delivered is False
    assert outcome.error_status == "send_failed"


def test_a_non_plan_is_refused():
    with pytest.raises(LoopbackError) as caught:
        deliver_reply({"chat_id": CHAT_ID}, send=ok_send())
    assert caught.value.status == REPLY_NOT_DELIVERABLE


# ------------------------------------------------------------------ result envelopes


def test_the_plan_result_carries_no_reply_text_and_no_chat_id():
    plan = ReplyPlan(chat_id=CHAT_ID, body="a private answer", content_type="text",
                     session_id="session-a", activity_id="activity-1")
    rendered = str(plan.to_result())
    assert CHAT_ID not in rendered
    assert "a private answer" not in rendered
    assert plan.to_result()["body_length"] == len("a private answer")


def test_the_outcome_result_carries_no_reply_text_and_no_chat_id():
    plan = ReplyPlan(chat_id=CHAT_ID, body="a private answer", content_type="text",
                     session_id="session-a", activity_id="activity-1")
    outcome = deliver_reply(plan, send=lambda arguments: {"status": "succeeded", "result": {"id": "m"}})
    rendered = str(outcome.to_result())
    assert CHAT_ID not in rendered
    assert "a private answer" not in rendered


def test_an_html_route_keeps_its_content_type():
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID, content_type="html"))
    plan = plan_reply_for(decision=decision(), router=router, render=lambda a: "<b>ok</b>")
    assert plan.content_type == "html"
    assert plan.to_send_arguments()["content_type"] == "html"

# ------------------------------------------------------------------ wired to the real send path


def test_the_planned_arguments_are_accepted_by_the_real_send_handler():
    """The loopback's output must be what the authorised operation actually takes.

    This is the integration claim the module is worth nothing without: the plan builds exactly
    three keys, and the real ``teams.send_chat_message`` handler posts them to the bound chat.
    """
    from urllib.parse import quote, urlsplit

    from tests.test_handlers_teams import HandlerGraphAdapter, invocation, open_context, serialized

    # The generated builder percent-encodes the chat id into the path.
    encoded = quote(CHAT_ID, safe="")
    adapter = HandlerGraphAdapter({("POST", f"/chats/{encoded}/messages"): {}})
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(decision=decision(), router=router, render=lambda a: "the answer")

    outcome = deliver_reply(
        plan,
        send=lambda arguments: invocation("teams.send_chat_message")(arguments, open_context(adapter)),
    )

    assert outcome.delivered is True
    request = adapter.requests[0]
    assert urlsplit(request.url).path == f"/chats/{encoded}/messages"
    assert serialized(request) == {"body": {"content": "the answer", "contentType": "text"}}


def test_the_loopback_never_posts_to_a_chat_the_binding_did_not_name():
    """A hostile body is present and the request still lands on the bound chat only."""
    from urllib.parse import quote, urlsplit

    from tests.test_handlers_teams import HandlerGraphAdapter, invocation, open_context

    hostile_chat = "19:attacker@thread.v2"
    encoded = quote(CHAT_ID, safe="")
    adapter = HandlerGraphAdapter({("POST", f"/chats/{encoded}/messages"): {}})
    router = ReplyRouter()
    router.bind(IDENTITY, ReplyRoute(chat_id=CHAT_ID))
    plan = plan_reply_for(
        decision=decision(model=activity(conversation=hostile_chat)),
        router=router,
        render=lambda a: "ok",
    )

    outcome = deliver_reply(
        plan,
        send=lambda arguments: invocation("teams.send_chat_message")(arguments, open_context(adapter)),
    )

    assert outcome.delivered is True
    paths = [urlsplit(request.url).path for request in adapter.requests]
    assert paths == [f"/chats/{encoded}/messages"]
    assert quote(hostile_chat, safe="") not in " ".join(paths)
