"""Inbound Teams transport boundary: authenticated Bot Framework Activities.

Why this module exists
----------------------
Every Teams module so far is *outbound*: this plugin calls Microsoft Graph, and Graph is the data
and identity domain. Conversational Teams is the other direction. A message arrives as a Bot
Framework **Activity** pushed to a webhook, and the plugin has to decide, on the strength of an
untrusted HTTP request, *which Hermes session should see it*.

That decision is the whole risk surface, so this module owns only the boundary and owns it
completely:

1. **The token is verified before anything else is read.** The request body is untrusted input;
   a claim inside it is not evidence of anything until the signature proves Microsoft sent it.
2. **Only known fields survive.** An Activity is projected onto a closed set of fields rather
   than passed through, so an unexpected key cannot ride along into a session.
3. **Identities are correlated, never trusted.** The external identity (tenant, AAD object id,
   conversation id) is mapped to a Hermes session by an explicit store. An unmapped sender is
   refused: assuming a session for an unknown identity is how one tenant's message lands in
   another tenant's conversation.
4. **Deliveries are deduplicated.** Bot Framework retries an Activity until it gets a 2xx. The
   same activity id must not reach a session twice.

Separation from Graph, per the project's own architecture rule
--------------------------------------------------------------
This module never builds a Graph client and never sends a message. It answers one question --
*"is this a genuine Activity from a known identity, and have I already delivered it?"* -- and
hands the answer to the caller. The outbound reply is a separate, separately-authorised step
(`teams.send_chat_message`). Keeping the two apart is what stops an inbound webhook from
becoming an unapproved outbound path.

What this module deliberately does NOT do
-----------------------------------------
It does not run an HTTP server, register a route, or hold a socket. It is the decision layer a
host calls; wiring a listener is the host's business and is the one part of this boundary that
cannot be verified offline.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Mapping, Sequence

from .errors import MESSAGES, GraphError

#: The only Activity types this boundary accepts. ``message`` is what a user typed; the
#: lifecycle events (``conversationUpdate``, ``typing``, ``messageReaction``) are how a session
#: learns it exists and is still alive. Anything else is refused rather than ignored silently.
ACCEPTED_ACTIVITY_TYPES = frozenset({"message", "conversationUpdate", "typing", "messageReaction"})

#: The token algorithm. Pinned to RS256 by name: accepting an algorithm named by the token
#: itself is the classic algorithm-confusion attack, and Bot Framework only ever signs RS256.
ACCEPTED_ALGORITHM = "RS256"

#: Bot Framework's own issuers. Both are checked; a token from anywhere else is refused.
BOT_FRAMEWORK_ISSUERS = ("https://api.botframework.com",)

#: Clock skew allowance when checking ``exp``/``nbf``, in seconds. Tokens live 5 minutes and
#: clocks are not perfectly aligned; a small explicit window beats a silent tolerance.
DEFAULT_LEEWAY_SECONDS = 300

#: How many recently-delivered activity ids are remembered for deduplication.
DEFAULT_DEDUPE_CAPACITY = 4096

INVALID_ACTIVITY = "invalid_activity"
UNSUPPORTED_ACTIVITY_TYPE = "unsupported_activity_type"
INVALID_TOKEN = "invalid_token"
TOKEN_EXPIRED = "token_expired"
TOKEN_AUDIENCE_MISMATCH = "token_audience_mismatch"
TOKEN_ISSUER_MISMATCH = "token_issuer_mismatch"
TOKEN_ALGORITHM_NOT_ALLOWED = "token_algorithm_not_allowed"
UNKNOWN_IDENTITY = "unknown_identity"
IDENTITY_BOUND_TO_ANOTHER_SESSION = "identity_bound_to_another_session"

_ACTIONS = {
    INVALID_ACTIVITY: "send a well-formed Bot Framework activity",
    UNSUPPORTED_ACTIVITY_TYPE: "only message and conversation lifecycle activities are accepted",
    INVALID_TOKEN: "send a bot framework activity with a valid bearer token",
    TOKEN_EXPIRED: "the activity token is stale; bot framework will retry with a fresh one",
    TOKEN_AUDIENCE_MISMATCH: "the token was not issued for this bot registration",
    TOKEN_ISSUER_MISMATCH: "the token was not issued by bot framework",
    TOKEN_ALGORITHM_NOT_ALLOWED: "the token is not signed with the accepted algorithm",
    UNKNOWN_IDENTITY: "register the external identity against a hermes session first",
    IDENTITY_BOUND_TO_ANOTHER_SESSION: "unbind the identity before rebinding it",
}


class InboundTeamsError(GraphError):
    """A sanitized inbound-boundary failure carrying an explicit, machine-readable ``status``.

    The rendered result never carries the token, the raw body, or an identity value: an inbound
    path that echoes what it rejected is an exfiltration path.
    """

    def __init__(
        self,
        category: str,
        *,
        status: str,
        action: str | None = None,
        activity_id: str | None = None,
    ):
        super().__init__(category, MESSAGES[category])
        self.status = str(status)
        self.action = action if action is not None else _ACTIONS.get(self.status)
        self.activity_id = activity_id

    def to_result(self) -> dict:
        payload = super().to_result()
        payload["status"] = self.status
        if self.action:
            payload["action"] = self.action
        if self.activity_id:
            payload["activity_id"] = self.activity_id
        return payload


# --------------------------------------------------------------------------- activity envelope


@dataclass(frozen=True)
class Activity:
    """The closed projection of one Bot Framework Activity.

    Only these fields survive. ``raw`` is intentionally absent: keeping the original body would
    mean every later log, result or error could leak it.
    """

    type: str
    activity_id: str | None
    timestamp: str | None
    text: str | None
    channel_id: str | None
    tenant_id: str | None
    sender_id: str | None
    sender_name: str | None
    recipient_id: str | None
    conversation_id: str | None
    conversation_type: str | None
    service_url: str | None
    locale: str | None
    members_added: tuple[str, ...] = field(default_factory=tuple)


def _text(value: Any) -> str | None:
    """A bounded, control-character-free string, or ``None``.

    Identity and text values reach a session and a log, so a raw control character or an
    unbounded value is refused rather than carried.
    """
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    if not stripped or len(stripped) > 4000:
        return None
    if any(ord(character) < 0x20 or ord(character) == 0x7F for character in stripped):
        return None
    return stripped


def _nested(payload: Mapping, *path: str) -> Mapping | None:
    current: Any = payload
    for name in path:
        if not isinstance(current, Mapping):
            return None
        current = current.get(name)
    return current if isinstance(current, Mapping) else None


def parse_activity(payload: Any) -> Activity:
    """Project an untrusted Activity onto the closed field set, or refuse it.

    Called *after* the token is verified: this function still validates everything it reads,
    because a correctly-signed Activity is not a well-formed one.
    """
    if not isinstance(payload, Mapping):
        raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)

    activity_type = _text(payload.get("type"))
    if activity_type is None:
        raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
    if activity_type not in ACCEPTED_ACTIVITY_TYPES:
        raise InboundTeamsError("validation_error", status=UNSUPPORTED_ACTIVITY_TYPE)

    conversation = _nested(payload, "conversation")
    sender = _nested(payload, "from")
    recipient = _nested(payload, "recipient")
    channel_data = _nested(payload, "channelData") or {}
    tenant = _nested(channel_data, "tenant")

    members: list[str] = []
    if activity_type == "conversationUpdate":
        for entry in payload.get("membersAdded") or ():
            member_id = _text((entry or {}).get("id")) if isinstance(entry, Mapping) else None
            if member_id:
                members.append(member_id)

    return Activity(
        type=activity_type,
        activity_id=_text(payload.get("id")),
        timestamp=_text(payload.get("timestamp")),
        text=_text(payload.get("text")),
        channel_id=_text(payload.get("channelId")),
        tenant_id=_text(tenant.get("id")) if tenant else None,
        sender_id=_text(sender.get("aadObjectId")) or _text(sender.get("id")) if sender else None,
        sender_name=_text(sender.get("name")) if sender else None,
        recipient_id=_text(recipient.get("id")) if recipient else None,
        conversation_id=_text(conversation.get("id")) if conversation else None,
        conversation_type=_text(conversation.get("conversationType")) if conversation else None,
        service_url=_text(payload.get("serviceUrl")),
        locale=_text(payload.get("locale")),
        members_added=tuple(members),
    )


# --------------------------------------------------------------------------- token verification

VerifySignature = Callable[[str, Sequence[Mapping], str, str | None], Mapping]

#: Failures that mean *we* could not check the token rather than that the token is bad. Kept
#: deliberately small and unambiguous: anything about the token itself stays an auth refusal,
#: and the specific "the token is not valid" family is recognised by PyJWT inside the seam.
_OPERATIONAL_VERIFICATION_FAILURES = (
    ConnectionError,
    TimeoutError,
    OSError,
    ImportError,
    MemoryError,
)


def _claims_via_pyjwt(token: str, keys: Sequence[Mapping], algorithm: str, audience: str | None) -> Mapping:
    """The production signature seam: real ``PyJWT`` verification against published JWKS.

    ``PyJWT`` is asked for the claims it verified. ``verify_exp``/``verify_aud``/``verify_iss``
    are left **off here on purpose**: this function only proves the signature, and the separate
    claim checks below are what enforce audience, issuer and lifetime under one explicit clock
    and one explicit leeway.
    """
    import jwt
    from jwt import PyJWKSet
    from jwt.exceptions import PyJWKError

    try:
        key_set = PyJWKSet.from_dict({"keys": list(keys)})
    except Exception as exc:  # noqa: BLE001 - the key set itself is unusable, not the token
        raise InboundTeamsError("configuration_error", status=INVALID_TOKEN) from exc
    header = jwt.get_unverified_header(token)
    key = next(
        (
            candidate
            for candidate in key_set.keys
            if candidate.key_id == header.get("kid")
        ),
        None,
    )
    if key is None:
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN)
    try:
        return jwt.decode(
            token,
            key.key,
            algorithms=[algorithm],
            audience=audience,
            options={"verify_exp": False, "verify_aud": False, "verify_iss": False},
        )
    except PyJWKError as exc:
        # The key could not be used at all -- this says nothing about the presented token.
        raise InboundTeamsError("configuration_error", status=INVALID_TOKEN) from exc


def verify_activity_token(
    token: Any,
    *,
    audience: str,
    keys: Iterable[Mapping] | None = None,
    verify_signature: VerifySignature | None = None,
    issuers: Sequence[str] = BOT_FRAMEWORK_ISSUERS,
    algorithm: str = ACCEPTED_ALGORITHM,
    leeway_seconds: int = DEFAULT_LEEWAY_SECONDS,
    now: float | None = None,
) -> Mapping:
    """Verify a Bot Framework bearer token and return its claims, or refuse it.

    The checks run in this order, and each one refuses rather than continues:

    1. **Signature** against Microsoft's published keys -- with the algorithm *pinned* to RS256,
       so a token cannot choose the algorithm that validates it.
    2. **Audience** -- the token must be issued for *this* bot registration. A token minted for
       another bot is a valid Microsoft signature and still not for us.
    3. **Issuer** -- Bot Framework only.
    4. **Lifetime** -- against one explicit clock and one explicit leeway, not the library's
       implicit defaults.

    ``verify_signature`` is the seam that makes this testable offline; production leaves it
    unset and gets real JWKS verification. **Fetching and caching the published key set is the
    caller's job** -- those keys rotate, and a boundary that fetched them inline would put a
    network round trip inside its own verification path. They are published at
    ``https://login.botframework.com/v1/.well-known/openidconfiguration``.
    """
    if not isinstance(token, str) or not token.strip() or len(token) > 16_384:
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN)
    claims_verifier = verify_signature
    if claims_verifier is None:
        if keys is None:
            raise InboundTeamsError("configuration_error", status=INVALID_TOKEN)
        claims_verifier = _claims_via_pyjwt

    try:
        header_algorithm, header = _token_header(token)
    except Exception:  # noqa: BLE001 - an unreadable token is an invalid token
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN) from None

    if header_algorithm != algorithm:
        # Refused before the signature is even attempted: this is the algorithm-confusion guard.
        raise InboundTeamsError("authentication_required", status=TOKEN_ALGORITHM_NOT_ALLOWED)

    try:
        claims = claims_verifier(token, tuple(keys or ()), algorithm, audience)
    except InboundTeamsError:
        raise
    except _OPERATIONAL_VERIFICATION_FAILURES:
        # A broken key set or an unreachable key endpoint is our problem, not the token's. It is
        # reported as a configuration error so the caller can retry after fixing it: answering
        # "unauthorized" would tell Bot Framework to stop retrying and drop a legitimate message
        # with nothing to alert on.
        raise InboundTeamsError("configuration_error", status=INVALID_TOKEN) from None
    except Exception:  # noqa: BLE001 - every signature failure is the same sanitized refusal
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN) from None

    if not isinstance(claims, Mapping):
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN)

    token_audience = claims.get("aud")
    if isinstance(token_audience, (list, tuple)):
        if audience not in token_audience:
            raise InboundTeamsError("authentication_required", status=TOKEN_AUDIENCE_MISMATCH)
    elif token_audience != audience:
        raise InboundTeamsError("authentication_required", status=TOKEN_AUDIENCE_MISMATCH)

    issuer = claims.get("iss")
    if issuer not in tuple(issuers):
        raise InboundTeamsError("authentication_required", status=TOKEN_ISSUER_MISMATCH)

    current = time.time() if now is None else now
    expires = claims.get("exp")
    if isinstance(expires, (int, float)) and current > float(expires) + leeway_seconds:
        raise InboundTeamsError("authentication_required", status=TOKEN_EXPIRED)
    not_before = claims.get("nbf")
    if isinstance(not_before, (int, float)) and current + leeway_seconds < float(not_before):
        raise InboundTeamsError("authentication_required", status=TOKEN_EXPIRED)
    return claims


def _token_header(token: str) -> tuple[str | None, Mapping]:
    """Read the (unverified) header, which is the only place the algorithm is declared."""
    import jwt

    header = jwt.get_unverified_header(token)
    if not isinstance(header, Mapping):
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN)
    algorithm = header.get("alg")
    return (algorithm if isinstance(algorithm, str) else None), header


def authorization_bearer(header_value: Any) -> str | None:
    """Extract a bearer token from an ``Authorization`` header value, or ``None``.

    Only the exact ``Bearer <token>`` form is accepted; a scheme mismatch or a multi-value
    header is refused rather than searched for something token-shaped.
    """
    if not isinstance(header_value, str):
        return None
    scheme, separator, value = header_value.partition(" ")
    if not separator or scheme.lower() != "bearer":
        return None
    token = value.strip()
    return token or None


# --------------------------------------------------------------------------- correlation


@dataclass(frozen=True)
class Identity:
    """The external identity a session is bound to.

    ``tenant_id`` is part of the key and not metadata: two tenants can present the same
    conversation shape, and a binding that ignored the tenant would let one tenant's message
    into another tenant's session.
    """

    tenant_id: str
    sender_id: str
    conversation_id: str

    def as_key(self) -> tuple[str, str, str]:
        return (self.tenant_id, self.sender_id, self.conversation_id)


def identity_of(activity: Activity) -> Identity:
    """The correlation key of an Activity, or a refusal when it cannot be formed."""
    if not activity.tenant_id or not activity.sender_id or not activity.conversation_id:
        raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
    return Identity(activity.tenant_id, activity.sender_id, activity.conversation_id)


class ConversationRegistry:
    """Binds external identities to Hermes sessions, and deduplicates deliveries.

    Both pieces of state are process-local and bounded, and the module says so out loud rather
    than implying a durability it does not have: a multi-replica deployment needs a shared
    store behind this same interface, and a restart forgets bindings by design.
    """

    def __init__(self, *, dedupe_capacity: int = DEFAULT_DEDUPE_CAPACITY):
        if isinstance(dedupe_capacity, bool) or not isinstance(dedupe_capacity, int) or dedupe_capacity < 1:
            raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
        self._capacity = dedupe_capacity
        self._bindings: dict[tuple[str, str, str], str] = {}
        self._sessions: dict[str, set[tuple[str, str, str]]] = {}
        self._delivered: dict[str, None] = {}

    # ---------------------------------------------------------------- bindings

    def bind(self, identity: Identity, session_id: str) -> None:
        """Bind an external identity to a Hermes session.

        Rebinding to a *different* session is refused: an identity that silently changed hands
        is exactly the confusion this store exists to prevent. Re-binding to the same session is
        idempotent, because a restart or a retry must not be an error.
        """
        if not isinstance(session_id, str) or not session_id.strip():
            raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
        key = identity.as_key()
        existing = self._bindings.get(key)
        if existing is not None and existing != session_id:
            raise InboundTeamsError("validation_error", status=IDENTITY_BOUND_TO_ANOTHER_SESSION)
        self._bindings[key] = session_id
        self._sessions.setdefault(session_id, set()).add(key)

    def unbind(self, identity: Identity) -> bool:
        """Remove a binding. Returns whether one existed."""
        key = identity.as_key()
        session_id = self._bindings.pop(key, None)
        if session_id is None:
            return False
        bound = self._sessions.get(session_id)
        if bound is not None:
            bound.discard(key)
            if not bound:
                self._sessions.pop(session_id, None)
        return True

    def session_for(self, identity: Identity) -> str:
        """The session bound to an identity, or a refusal when there is none.

        There is no default session and no "first available" session: an unbound sender is
        refused, because guessing is how a message reaches a conversation it does not belong to.
        """
        session_id = self._bindings.get(identity.as_key())
        if session_id is None:
            raise InboundTeamsError("validation_error", status=UNKNOWN_IDENTITY)
        return session_id

    def identities_for(self, session_id: str) -> tuple[Identity, ...]:
        return tuple(
            Identity(*key) for key in sorted(self._sessions.get(session_id, ()))
        )

    # ---------------------------------------------------------------- deliveries

    def already_delivered(self, activity_id: str) -> bool:
        if not isinstance(activity_id, str) or not activity_id:
            raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
        return activity_id in self._delivered

    def mark_delivered(self, activity_id: str) -> None:
        """Record a delivery, evicting the oldest entry past the capacity.

        Bounded on purpose: retries arrive within minutes, so a small window catches every real
        duplicate, and an unbounded set would be a memory exhaustion path fed by the network.
        """
        if not isinstance(activity_id, str) or not activity_id:
            raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)
        self._delivered.pop(activity_id, None)
        self._delivered[activity_id] = None
        while len(self._delivered) > self._capacity:
            oldest = next(iter(self._delivered))
            self._delivered.pop(oldest, None)


@dataclass(frozen=True)
class InboundDecision:
    """The boundary's answer: which session the Activity belongs to, and what it carried."""

    session_id: str
    identity: Identity
    activity: Activity
    duplicate: bool = False

    def to_result(self) -> dict:
        """The rendered envelope.

        Identity values are summarized rather than echoed: a result that carries a raw tenant or
        sender id would put them in every transcript that logs a decision.
        """
        return {
            "session_bound": bool(self.session_id),
            "activity_type": self.activity.type,
            "activity_id": self.activity.activity_id,
            "duplicate": self.duplicate,
            "has_text": self.activity.text is not None,
            "conversation_type": self.activity.conversation_type,
            "conversation_ref_present": self.activity.service_url is not None,
        }


def admit_activity(
    *,
    authorization: Any,
    payload: Any,
    registry: ConversationRegistry,
    audience: str,
    keys: Iterable[Mapping] | None = None,
    verify_signature: VerifySignature | None = None,
    issuers: Sequence[str] = BOT_FRAMEWORK_ISSUERS,
    leeway_seconds: int = DEFAULT_LEEWAY_SECONDS,
    now: float | None = None,
) -> InboundDecision:
    """The single entry point: authenticate, validate, correlate, deduplicate.

    Order matters and is not negotiable:

    1. the token is verified -- an unauthenticated body is never parsed for meaning;
    2. the Activity is projected onto the closed field set;
    3. the identity is correlated to a bound session, or refused;
    4. a repeat delivery is reported as a duplicate, not delivered again.

    A duplicate is returned as a decision with ``duplicate=True`` rather than raised: the correct
    HTTP answer for a retried Activity is a success, and the caller needs to know it was a
    retry so it does not hand the same text to the session twice.
    """
    token = authorization_bearer(authorization)
    if token is None:
        raise InboundTeamsError("authentication_required", status=INVALID_TOKEN)
    verify_activity_token(
        token,
        audience=audience,
        keys=keys,
        verify_signature=verify_signature,
        issuers=issuers,
        leeway_seconds=leeway_seconds,
        now=now,
    )
    activity = parse_activity(payload)
    identity = identity_of(activity)
    session_id = registry.session_for(identity)

    if activity.activity_id is None:
        # An Activity with no id cannot be deduplicated, so it is refused rather than delivered
        # twice on a retry: "cannot tell" must not become "deliver anyway".
        raise InboundTeamsError("validation_error", status=INVALID_ACTIVITY)

    if registry.already_delivered(activity.activity_id):
        return InboundDecision(session_id, identity, activity, duplicate=True)

    registry.mark_delivered(activity.activity_id)
    return InboundDecision(session_id, identity, activity, duplicate=False)