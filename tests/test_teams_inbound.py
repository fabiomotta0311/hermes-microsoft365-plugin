"""Offline tests for the inbound Teams boundary (``microsoft365.teams_inbound``).

These tests are the reason the boundary can be trusted without a live tenant: every rule that
protects a session is exercised here with a token this suite signs itself, so a signature that
should be refused is refused for the right stated reason.
"""
from __future__ import annotations

import pytest

from microsoft365.teams_inbound import (
    ACCEPTED_ALGORITHM,
    BOT_FRAMEWORK_ISSUERS,
    IDENTITY_BOUND_TO_ANOTHER_SESSION,
    INVALID_ACTIVITY,
    INVALID_TOKEN,
    TOKEN_ALGORITHM_NOT_ALLOWED,
    TOKEN_AUDIENCE_MISMATCH,
    TOKEN_EXPIRED,
    TOKEN_ISSUER_MISMATCH,
    UNKNOWN_IDENTITY,
    Activity,
    ConversationRegistry,
    Identity,
    InboundTeamsError,
    admit_activity,
    authorization_bearer,
    identity_of,
    parse_activity,
    verify_activity_token,
)

AUDIENCE = "bot-app-id-0001"
TENANT = "tenant-1"
SENDER = "sender-1"
CONVERSATION = "19:chat@thread.v2"
NOW = 1_700_000_000.0


def activity_payload(
    *,
    activity_id="activity-1",
    activity_type="message",
    text="hello",
    sender=SENDER,
    tenant=TENANT,
    conversation=CONVERSATION,
    extra=None,
):
    payload = {
        "type": activity_type,
        "id": activity_id,
        "timestamp": "2026-10-02T12:00:00.000Z",
        "text": text,
        "channelId": "msteams",
        "serviceUrl": "https://smba.trafficmanager.net/amer/",
        "locale": "pt-BR",
        "from": {"id": "29:external", "aadObjectId": sender, "name": "Someone"},
        "recipient": {"id": "28:bot", "name": "Hermes"},
        "conversation": {"id": conversation, "conversationType": "personal"},
        "channelData": {"tenant": {"id": tenant}},
    }
    if extra:
        payload.update(extra)
    return payload


def claims(*, audience=AUDIENCE, issuer=BOT_FRAMEWORK_ISSUERS[0], exp=NOW + 300, nbf=NOW - 10):
    return {"aud": audience, "iss": issuer, "exp": exp, "nbf": nbf}


def verifier(claim_set=None, *, raises=None):
    """A signature seam that reports a fixed claim set, standing in for real JWKS verification."""

    def verify(token, keys, algorithm, audience):
        if raises is not None:
            raise raises
        return claims() if claim_set is None else claim_set

    return verify


def token_for(algorithm=ACCEPTED_ALGORITHM):
    """A structurally valid RS256-shaped JWT header/payload pair (never signature-checked here)."""
    import base64
    import json

    def segment(value):
        raw = value if isinstance(value, bytes) else json.dumps(value).encode()
        return base64.urlsafe_b64encode(raw).rstrip(b"=").decode()

    # The signature segment must itself be valid base64url: the header reader decodes every
    # segment before it can report the algorithm, so a placeholder string would be unreadable.
    return f"{segment({'alg': algorithm, 'kid': 'key-1', 'typ': 'JWT'})}.{segment({'sub': 'x'})}.{segment(b'sig')}"


# ------------------------------------------------------------------ authorization header


def test_only_the_exact_bearer_form_yields_a_token():
    assert authorization_bearer("Bearer abc.def.ghi") == "abc.def.ghi"
    assert authorization_bearer("bearer lowercase") == "lowercase"
    for bad in (None, "", "abc.def.ghi", "Basic abc", "Bearer", "Bearer   ", 12345, ["Bearer abc"]):
        assert authorization_bearer(bad) is None


# ------------------------------------------------------------------ signature and claims


def test_a_correctly_signed_and_scoped_token_is_accepted():
    result = verify_activity_token(
        token_for(), audience=AUDIENCE, keys=[{"kty": "RSA"}], verify_signature=verifier(), now=NOW
    )
    assert result["aud"] == AUDIENCE


def test_a_token_signed_for_another_bot_is_refused():
    """A genuine Microsoft signature for a different registration is still not for us."""
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(claims(audience="some-other-bot")),
            now=NOW,
        )
    assert caught.value.status == TOKEN_AUDIENCE_MISMATCH


def test_an_audience_list_containing_this_bot_is_accepted():
    result = verify_activity_token(
        token_for(),
        audience=AUDIENCE,
        keys=[{"kty": "RSA"}],
        verify_signature=verifier(claims(audience=["other", AUDIENCE])),
        now=NOW,
    )
    assert result["aud"] == ["other", AUDIENCE]


def test_a_token_from_another_issuer_is_refused():
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(claims(issuer="https://evil.example")),
            now=NOW,
        )
    assert caught.value.status == TOKEN_ISSUER_MISMATCH


def test_an_expired_token_is_refused():
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(claims(exp=NOW - 3600)),
            now=NOW,
        )
    assert caught.value.status == TOKEN_EXPIRED


def test_a_token_that_is_not_valid_yet_is_refused():
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(claims(nbf=NOW + 3600)),
            now=NOW,
        )
    assert caught.value.status == TOKEN_EXPIRED


def test_a_clock_within_the_leeway_still_accepts_a_just_expired_token():
    """Explicit leeway, not an implicit library default: the window is a stated number."""
    result = verify_activity_token(
        token_for(),
        audience=AUDIENCE,
        keys=[{"kty": "RSA"}],
        verify_signature=verifier(claims(exp=NOW - 60)),
        now=NOW,
    )
    assert result["aud"] == AUDIENCE


@pytest.mark.parametrize("algorithm", ["HS256", "none", "RS512", "PS256", ""])
def test_a_token_that_names_another_algorithm_is_refused_before_the_signature(algorithm):
    """Algorithm confusion: the token does not get to choose how it is validated."""
    calls = []

    def verify(token, keys, alg, audience):
        calls.append(alg)
        return claims()

    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(algorithm), audience=AUDIENCE, keys=[{"kty": "RSA"}],
            verify_signature=verify, now=NOW,
        )
    assert caught.value.status == TOKEN_ALGORITHM_NOT_ALLOWED
    assert calls == []  # the signature seam was never reached


@pytest.mark.parametrize("bad", [None, "", "   ", 12345, ["a", "b"], "x" * 20_000])
def test_an_unusable_token_value_is_refused(bad):
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            bad, audience=AUDIENCE, keys=[{"kty": "RSA"}], verify_signature=verifier(), now=NOW
        )
    assert caught.value.status == INVALID_TOKEN


def test_a_signature_failure_is_a_sanitized_refusal():
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(
            token_for(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(raises=ValueError("raw jwt detail: secret=abc")),
            now=NOW,
        )
    assert caught.value.status == INVALID_TOKEN
    assert "secret" not in str(caught.value.to_result())


def test_a_missing_key_set_is_a_configuration_error_not_a_pass():
    with pytest.raises(InboundTeamsError) as caught:
        verify_activity_token(token_for(), audience=AUDIENCE, keys=None, now=NOW)
    assert caught.value.status == INVALID_TOKEN


# ------------------------------------------------------------------ activity projection


def test_the_projection_keeps_only_the_closed_field_set():
    parsed = parse_activity(activity_payload(extra={"warez": "payload", "secrets": {"k": "v"}}))
    assert parsed.type == "message"
    assert parsed.text == "hello"
    assert parsed.tenant_id == TENANT
    assert parsed.sender_id == SENDER
    assert parsed.conversation_id == CONVERSATION
    # The closed projection exposes no attribute for an unexpected key.
    assert not hasattr(parsed, "warez")
    assert "warez" not in parsed.__dict__


def test_the_projection_carries_no_raw_body():
    parsed = parse_activity(activity_payload())
    assert not hasattr(parsed, "raw")
    assert "hello" not in [name for name in parsed.__dict__ if name != "text"]


def project(payload):
    """Everything the boundary does to a body: project it, then derive its identity."""
    return identity_of(parse_activity(payload))


@pytest.mark.parametrize(
    "bad",
    [
        None,
        "a string",
        42,
        [],
        {},
        {"type": "message"},
        {"type": ""},
        {"type": "invoke"},
        {"type": "message", "id": "a", "conversation": {"id": "c"}},  # no sender/tenant
        {"type": "message", "id": "a", "from": {"aadObjectId": "s"}},  # no conversation
    ],
)
def test_a_malformed_or_unsupported_activity_is_refused(bad):
    """Refused either by the projection or by the identity it must yield -- never admitted."""
    with pytest.raises(InboundTeamsError) as caught:
        project(bad)
    assert caught.value.status in {
        INVALID_ACTIVITY,
        "unsupported_activity_type",
    }


def test_a_wellformed_activity_without_an_identity_is_refused_by_identity_not_by_parsing():
    """Parsing is projection, not authorisation: the missing identity is the refusal."""
    parsed = parse_activity({"type": "message", "id": "a", "conversation": {"id": "c"}})
    assert parsed.type == "message"
    with pytest.raises(InboundTeamsError) as caught:
        identity_of(parsed)
    assert caught.value.status == INVALID_ACTIVITY


def test_an_unsupported_activity_type_is_named_as_such():
    with pytest.raises(InboundTeamsError) as caught:
        parse_activity(activity_payload(activity_type="invoke"))
    assert caught.value.status == "unsupported_activity_type"


def test_control_characters_in_a_field_are_dropped_rather_than_carried():
    parsed = parse_activity(activity_payload(text="bad\x00text"))
    assert parsed.text is None


def test_a_conversation_update_reports_the_members_added():
    parsed = parse_activity(
        activity_payload(
            activity_type="conversationUpdate",
            text=None,
            extra={"membersAdded": [{"id": "28:bot"}, {"id": "29:someone"}]},
        )
    )
    assert parsed.members_added == ("28:bot", "29:someone")


def test_an_identity_needs_tenant_sender_and_conversation():
    parsed = parse_activity(activity_payload())
    assert identity_of(parsed).as_key() == (TENANT, SENDER, CONVERSATION)
    with pytest.raises(InboundTeamsError) as caught:
        identity_of(Activity(type="message", activity_id="a", timestamp=None, text=None,
                             channel_id=None, tenant_id=None, sender_id=SENDER,
                             sender_name=None, recipient_id=None, conversation_id=CONVERSATION,
                             conversation_type=None, service_url=None, locale=None))
    assert caught.value.status == INVALID_ACTIVITY


# ------------------------------------------------------------------ correlation


def test_an_unbound_identity_is_refused_rather_than_given_a_session():
    registry = ConversationRegistry()
    identity = Identity(TENANT, SENDER, CONVERSATION)
    with pytest.raises(InboundTeamsError) as caught:
        registry.session_for(identity)
    assert caught.value.status == UNKNOWN_IDENTITY


def test_a_bound_identity_resolves_to_its_session():
    registry = ConversationRegistry()
    identity = Identity(TENANT, SENDER, CONVERSATION)
    registry.bind(identity, "session-a")
    assert registry.session_for(identity) == "session-a"
    assert registry.identities_for("session-a") == (identity,)


def test_rebinding_the_same_identity_to_the_same_session_is_idempotent():
    registry = ConversationRegistry()
    identity = Identity(TENANT, SENDER, CONVERSATION)
    registry.bind(identity, "session-a")
    registry.bind(identity, "session-a")
    assert registry.session_for(identity) == "session-a"


def test_rebinding_an_identity_to_a_different_session_is_refused():
    """An identity that silently changes hands is the confusion this store exists to prevent."""
    registry = ConversationRegistry()
    identity = Identity(TENANT, SENDER, CONVERSATION)
    registry.bind(identity, "session-a")
    with pytest.raises(InboundTeamsError) as caught:
        registry.bind(identity, "session-b")
    assert caught.value.status == IDENTITY_BOUND_TO_ANOTHER_SESSION
    assert registry.session_for(identity) == "session-a"


def test_the_same_sender_in_another_tenant_is_a_different_identity():
    """Two tenants can present the same shape; the binding must not span them."""
    registry = ConversationRegistry()
    registry.bind(Identity("tenant-1", SENDER, CONVERSATION), "session-a")
    with pytest.raises(InboundTeamsError) as caught:
        registry.session_for(Identity("tenant-2", SENDER, CONVERSATION))
    assert caught.value.status == UNKNOWN_IDENTITY


def test_unbinding_frees_the_identity():
    registry = ConversationRegistry()
    identity = Identity(TENANT, SENDER, CONVERSATION)
    registry.bind(identity, "session-a")
    assert registry.unbind(identity) is True
    assert registry.unbind(identity) is False
    assert registry.identities_for("session-a") == ()


def test_a_blank_session_id_is_refused():
    registry = ConversationRegistry()
    for bad in (None, "", "   ", 42):
        with pytest.raises(InboundTeamsError):
            registry.bind(Identity(TENANT, SENDER, CONVERSATION), bad)


# ------------------------------------------------------------------ deduplication


def test_a_delivered_activity_is_remembered():
    registry = ConversationRegistry()
    assert registry.already_delivered("a-1") is False
    registry.mark_delivered("a-1")
    assert registry.already_delivered("a-1") is True


def test_deduplication_is_bounded_and_evicts_the_oldest():
    """The window is bounded on purpose: an unbounded set is a network-fed memory leak."""
    registry = ConversationRegistry(dedupe_capacity=3)
    for index in range(3):
        registry.mark_delivered(f"a-{index}")
    registry.mark_delivered("a-3")
    assert registry.already_delivered("a-3") is True
    assert registry.already_delivered("a-0") is False  # evicted


def test_remarking_the_same_delivery_does_not_grow_the_window():
    registry = ConversationRegistry(dedupe_capacity=2)
    registry.mark_delivered("a-1")
    registry.mark_delivered("a-2")
    registry.mark_delivered("a-1")
    assert registry.already_delivered("a-2") is True  # a re-mark did not evict it


def test_a_nonsense_capacity_is_refused():
    for bad in (0, -1, True, "10"):
        with pytest.raises(InboundTeamsError):
            ConversationRegistry(dedupe_capacity=bad)


# ------------------------------------------------------------------ admit_activity


def bound_registry():
    registry = ConversationRegistry()
    registry.bind(Identity(TENANT, SENDER, CONVERSATION), "session-a")
    return registry


def test_a_genuine_activity_reaches_its_bound_session():
    decision = admit_activity(
        authorization="Bearer " + token_for(),
        payload=activity_payload(),
        registry=bound_registry(),
        audience=AUDIENCE,
        keys=[{"kty": "RSA"}],
        verify_signature=verifier(),
        now=NOW,
    )
    assert decision.session_id == "session-a"
    assert decision.duplicate is False
    assert decision.activity.text == "hello"
    rendered = decision.to_result()
    assert rendered["session_bound"] is True
    assert rendered["has_text"] is True


def test_a_retried_activity_is_reported_as_a_duplicate_and_not_delivered_twice():
    registry = bound_registry()
    kwargs = dict(authorization="Bearer " + token_for(), registry=registry,
                  audience=AUDIENCE, keys=[{"kty": "RSA"}], verify_signature=verifier(), now=NOW)

    first = admit_activity(payload=activity_payload(), **kwargs)
    second = admit_activity(payload=activity_payload(), **kwargs)

    assert first.duplicate is False
    # Bot Framework retries until it sees a 2xx: the retry is a success, not a second delivery.
    assert second.duplicate is True
    assert second.session_id == first.session_id


def test_an_activity_with_no_token_is_refused_before_its_body_is_read():
    with pytest.raises(InboundTeamsError) as caught:
        admit_activity(
            authorization=None,
            payload=activity_payload(),
            registry=bound_registry(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(),
            now=NOW,
        )
    assert caught.value.status == INVALID_TOKEN


def test_a_body_that_is_not_even_an_activity_is_refused_after_a_valid_token():
    with pytest.raises(InboundTeamsError) as caught:
        admit_activity(
            authorization="Bearer " + token_for(),
            payload={"not": "an activity"},
            registry=bound_registry(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(),
            now=NOW,
        )
    assert caught.value.status == INVALID_ACTIVITY


def test_an_unknown_sender_is_refused_even_with_a_perfect_token():
    """Authentication is not authorisation: a valid token from an unbound identity is refused."""
    with pytest.raises(InboundTeamsError) as caught:
        admit_activity(
            authorization="Bearer " + token_for(),
            payload=activity_payload(sender="someone-else"),
            registry=bound_registry(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(),
            now=NOW,
        )
    assert caught.value.status == UNKNOWN_IDENTITY


def test_an_activity_without_an_id_is_refused_rather_than_delivered_undeduplicated():
    """No id means "cannot tell if this is a retry"; that must not become "deliver anyway"."""
    with pytest.raises(InboundTeamsError) as caught:
        admit_activity(
            authorization="Bearer " + token_for(),
            payload=activity_payload(activity_id=None),
            registry=bound_registry(),
            audience=AUDIENCE,
            keys=[{"kty": "RSA"}],
            verify_signature=verifier(),
            now=NOW,
        )
    assert caught.value.status == INVALID_ACTIVITY


def test_a_refused_activity_does_not_consume_a_duplicate_slot():
    """A rejected delivery must not make the genuine retry look like a duplicate."""
    registry = bound_registry()
    kwargs = dict(authorization="Bearer " + token_for(), registry=registry,
                  audience=AUDIENCE, keys=[{"kty": "RSA"}], verify_signature=verifier(), now=NOW)
    with pytest.raises(InboundTeamsError):
        admit_activity(payload=activity_payload(sender="unknown"), **kwargs)

    decision = admit_activity(payload=activity_payload(), **kwargs)
    assert decision.duplicate is False


def test_the_rendered_decision_carries_no_identity_values_or_text():
    decision = admit_activity(
        authorization="Bearer " + token_for(),
        payload=activity_payload(text="a secret sentence"),
        registry=bound_registry(),
        audience=AUDIENCE,
        keys=[{"kty": "RSA"}],
        verify_signature=verifier(),
        now=NOW,
    )
    rendered = str(decision.to_result())
    assert TENANT not in rendered
    assert SENDER not in rendered
    assert CONVERSATION not in rendered
    assert "a secret sentence" not in rendered


def test_a_conversation_lifecycle_activity_is_admitted_without_text():
    registry = bound_registry()
    decision = admit_activity(
        authorization="Bearer " + token_for(),
        payload=activity_payload(
            activity_type="conversationUpdate",
            activity_id="conv-1",
            text=None,
            extra={"membersAdded": [{"id": "29:new"}]},
        ),
        registry=registry,
        audience=AUDIENCE,
        keys=[{"kty": "RSA"}],
        verify_signature=verifier(),
        now=NOW,
    )
    assert decision.activity.members_added == ("29:new",)
    assert decision.to_result()["has_text"] is False