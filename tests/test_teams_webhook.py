"""The webhook's contract, exercised over WSGI and once over a real socket.

The last section is the point of the file: the claim that a Bot Framework endpoint "cannot be
verified offline" was wrong. A WSGI application runs on the standard library, so the endpoint is
started on a real port and called with ``http.client`` -- a genuine HTTP request over a genuine
socket, with no cloud service involved.
"""
from __future__ import annotations

import json
import threading
from wsgiref.simple_server import WSGIRequestHandler, make_server

import pytest

from microsoft365.teams_inbound import ConversationRegistry, Identity, InboundTeamsError
from microsoft365.teams_loopback import ReplyRoute, ReplyRouter, deliver_reply, plan_reply_for
from microsoft365.teams_webhook import (
    DEFAULT_WEBHOOK_PATH,
    MAX_ACTIVITY_BYTES,
    build_application,
)

from tests.test_teams_inbound import (
    AUDIENCE,
    CONVERSATION,
    NOW,
    SENDER,
    TENANT,
    activity_payload,
    token_for,
    verifier,
)
from tests.test_teams_loopback import CHAT_ID

BEARER = f"Bearer {token_for()}"


def identity() -> Identity:
    """The identity the fixture Activity projects to: the binding key the registry uses."""
    return Identity(TENANT, SENDER, CONVERSATION)


def registry_with_binding(*, session_id="session-1"):
    registry = ConversationRegistry()
    registry.bind(identity(), session_id)
    return registry


def routed(*, chat_id=CHAT_ID) -> ReplyRouter:
    """A router with a route bound for the fixture identity."""
    router = ReplyRouter()
    router.bind(identity(), ReplyRoute(chat_id=chat_id))
    return router


def application(
    *,
    registry=None,
    on_activity=None,
    reply=None,
    clock=lambda: NOW,
    path=DEFAULT_WEBHOOK_PATH,
    audience=AUDIENCE,
    verify_signature=None,
):
    return build_application(
        registry=registry if registry is not None else registry_with_binding(),
        audience=audience,
        verify_signature=verify_signature if verify_signature is not None else verifier(),
        path=path,
        on_activity=on_activity,
        reply=reply,
        clock=clock,
    )


def payload_bytes(payload=None) -> bytes:
    return json.dumps(payload if payload is not None else activity_payload()).encode("utf-8")


def call(
    app,
    *,
    method="POST",
    path=DEFAULT_WEBHOOK_PATH,
    body=None,
    authorization=BEARER,
    content_type="application/json",
    declare_length="auto",
    omit_length=False,
):
    """Drive the WSGI app directly and return (status_code, headers, parsed_body, raw_bytes)."""
    raw = payload_bytes() if body is None else body
    if omit_length:
        declare_length = None
    elif declare_length == "auto":
        declare_length = len(raw)
    environ = {
        "REQUEST_METHOD": method,
        "PATH_INFO": path,
        "CONTENT_TYPE": content_type,
        "wsgi.input": __import__("io").BytesIO(raw),
    }
    if authorization is not None:
        environ["HTTP_AUTHORIZATION"] = authorization
    if declare_length is not None:
        environ["CONTENT_LENGTH"] = str(declare_length)

    captured = {}

    def start_response(status, headers):
        captured["status"] = status
        captured["headers"] = dict(headers)

    chunks = app(environ, start_response)
    raw_body = b"".join(chunks)
    code = int(captured["status"].split(" ", 1)[0])
    try:
        parsed = json.loads(raw_body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        parsed = None
    return code, captured["headers"], parsed, raw_body


# ------------------------------------------------------------------ routing and method


def test_only_the_configured_path_is_served():
    code, _, body, _ = call(application(), path="/teams/activities/../admin")
    assert code == 404
    assert body == {"status": "not_found"}


def test_another_path_does_not_reveal_where_the_endpoint_is():
    code, _, body, _ = call(application(), path="/")
    assert code == 404
    assert body == {"status": "not_found"}


@pytest.mark.parametrize("method", ["GET", "PUT", "DELETE", "PATCH", "HEAD"])
def test_anything_but_post_is_refused_and_says_what_is_allowed(method):
    code, headers, body, _ = call(application(), method=method)
    assert code == 405
    assert headers["Allow"] == "POST"
    assert body == {"status": "method_not_allowed"}


def test_a_custom_path_is_honoured():
    code, _, _, _ = call(application(path="/hooks/msteams"), path="/hooks/msteams")
    assert code == 200


# ------------------------------------------------------------------ body handling


def test_a_missing_media_type_is_refused():
    code, _, body, _ = call(application(), content_type="")
    assert code == 415
    assert body == {"status": "unsupported_media_type"}


def test_a_non_json_media_type_is_refused():
    code, _, _, _ = call(application(), content_type="text/plain")
    assert code == 415


def test_a_json_media_type_with_a_charset_parameter_is_accepted():
    code, _, _, _ = call(application(), content_type="application/json; charset=utf-8")
    assert code == 200


def test_a_missing_content_length_is_refused_rather_than_read_until_eof():
    code, _, body, _ = call(application(), omit_length=True)
    assert code == 411
    assert body == {"status": "length_required"}


def test_a_non_numeric_content_length_is_refused():
    code, _, _, _ = call(application(), declare_length="not-a-number")
    assert code == 400


def test_a_negative_content_length_is_refused():
    code, _, _, _ = call(application(), declare_length=-1)
    assert code == 400


def test_an_oversized_declared_body_is_refused_before_it_is_read():
    code, _, body, _ = call(application(), declare_length=MAX_ACTIVITY_BYTES + 1)
    assert code == 413
    assert body == {"status": "payload_too_large"}


def test_an_understated_content_length_cannot_smuggle_an_oversized_body():
    """A header that lies is not a licence: the stream itself is read under the same bound."""
    oversized = b"x" * (MAX_ACTIVITY_BYTES + 10)
    code, _, body, _ = call(application(), body=oversized, declare_length=10)
    assert code in (400, 413)
    assert body["status"] in {"invalid_request", "payload_too_large"}


def test_reading_is_delimited_by_the_declared_length_not_by_the_stream():
    """The regression that would have deadlocked every real deployment.

    Under WSGI, ``wsgi.input`` may be the socket: asking for more bytes than the body contains
    does not return what is buffered, it waits for the peer to close. So the reader must ask for
    exactly ``Content-Length`` bytes. This stream would answer an unbounded request forever.
    """
    import io

    from microsoft365.teams_webhook import _read_declared

    class EndlessStream:
        def __init__(self):
            self.asked = []

        def read(self, size=-1):
            self.asked.append(size)
            return b"x" * (65536 if size < 0 else size)

    stream = EndlessStream()
    payload = _read_declared(stream, 10)
    assert payload == b"x" * 10
    assert stream.asked == [10], "the reader asked for more than the declared body length"


def test_a_body_shorter_than_its_declared_length_is_refused():
    from microsoft365.teams_webhook import _read_declared

    import pytest as _pytest

    from microsoft365.teams_webhook import WebhookError

    with _pytest.raises(WebhookError) as caught:
        _read_declared(__import__("io").BytesIO(b"short"), 100)
    assert caught.value.status_code == 400


def test_an_empty_body_is_refused():
    code, _, _, _ = call(application(), body=b"", declare_length=0)
    assert code == 400


def test_a_body_that_is_not_json_is_refused():
    code, _, body, _ = call(application(), body=b"not json at all")
    assert code == 400
    assert body == {"status": "invalid_request"}


def test_a_body_that_is_not_utf8_is_refused():
    code, _, _, _ = call(application(), body=b"\xff\xfe\x00bad")
    assert code == 400


def test_a_json_body_that_is_not_an_activity_is_refused():
    code, _, _, _ = call(application(), body=json.dumps([1, 2, 3]).encode())
    assert code == 400


# ------------------------------------------------------------------ authentication


def test_a_missing_authorization_header_is_refused():
    code, _, body, _ = call(application(), authorization=None)
    assert code == 401
    assert body == {"status": "unauthorized"}


@pytest.mark.parametrize(
    "header",
    ["", "Bearer", "Basic abc", "Bearer  two-spaces", "bearer lowercase-scheme", "Bearer a.b"],
)
def test_a_malformed_authorization_header_is_refused(header):
    code, _, _, _ = call(application(), authorization=header)
    assert code == 401


def test_every_distinguishable_token_failure_leaves_as_the_same_anonymous_401():
    """The reason a token was refused is exactly the feedback a forged token needs."""
    from microsoft365.teams_inbound import (
        TOKEN_ALGORITHM_NOT_ALLOWED,
        TOKEN_AUDIENCE_MISMATCH,
        TOKEN_EXPIRED,
        TOKEN_ISSUER_MISMATCH,
    )

    refusals = [
        InboundTeamsError("authentication_required", status=TOKEN_ALGORITHM_NOT_ALLOWED),
        InboundTeamsError("authentication_required", status=TOKEN_AUDIENCE_MISMATCH),
        InboundTeamsError("authentication_required", status=TOKEN_ISSUER_MISMATCH),
        InboundTeamsError("authentication_required", status=TOKEN_EXPIRED),
    ]
    bodies = []
    for refusal in refusals:
        app = application(verify_signature=verifier(raises=refusal))
        code, _, body, raw = call(app)
        assert code == 401
        bodies.append((json.dumps(body, sort_keys=True), raw))
    assert len({b for b, _ in bodies}) == 1, "the refusals are distinguishable from the response"


def test_a_token_failure_does_not_echo_the_token_back():
    app = application(verify_signature=verifier(raises=InboundTeamsError("authentication_required", status="invalid_token")))
    code, _, _, raw = call(app, authorization="Bearer super-secret-token-value")
    assert code == 401
    assert b"super-secret-token-value" not in raw


def test_a_configuration_failure_is_a_500_so_a_fixed_deployment_can_be_retried():
    """A boundary that cannot check a token must not tell the caller the token was bad."""
    app = application(
        verify_signature=verifier(
            raises=InboundTeamsError("configuration_error", status="invalid_token")
        )
    )
    code, _, body, _ = call(app)
    assert code == 500
    assert body == {"status": "internal_error"}


def test_an_unreachable_key_endpoint_is_retryable_rather_than_an_auth_refusal():
    """The published JWKS lives over the network; the network being down is not a bad token."""
    def unreachable(token, keys, algorithm, audience):
        raise ConnectionError("key endpoint unreachable")

    code, _, body, _ = call(application(verify_signature=unreachable))
    assert code == 500, "an unreachable key set must be retryable, not a permanent 401"
    assert body == {"status": "internal_error"}


def test_a_failing_session_handoff_does_not_leak_a_traceback():
    def explode(decision):
        raise RuntimeError("internal detail that must not travel")

    code, _, body, raw = call(application(on_activity=explode))
    assert code == 500
    assert body == {"status": "internal_error"}
    assert b"internal detail" not in raw


def test_a_verification_seam_failure_never_leaks_what_it_was_given():
    def explode(token, keys, algorithm, audience):
        raise RuntimeError("secret=abc raw jwt detail")

    code, _, body, raw = call(application(verify_signature=explode))
    assert code == 401, "a seam that raises is refusing the token: the seam contract"
    assert body == {"status": "unauthorized"}
    assert b"secret=abc" not in raw
    assert b"raw jwt detail" not in raw


# ------------------------------------------------------------------ delivery


def test_an_unbound_identity_is_refused_without_reaching_the_session():
    delivered = []
    app = application(
        registry=ConversationRegistry(),
        on_activity=delivered.append,
    )
    code, _, _, _ = call(app)
    assert code == 400
    assert delivered == []


def test_an_admitted_activity_reaches_the_session_before_any_reply():
    order = []
    app = application(
        on_activity=lambda decision: order.append(("session", decision.session_id)),
        reply=lambda decision: order.append(("reply", decision.session_id)) or None,
    )
    code, _, body, _ = call(app)
    assert code == 200
    assert body == {"status": "accepted"}
    assert [step for step, _ in order] == ["session", "reply"]


def test_a_duplicate_is_acknowledged_and_never_delivered_twice():
    seen = []
    app = application(on_activity=seen.append)
    first, _, _, _ = call(app)
    second, _, body, _ = call(app)
    assert first == 200
    assert second == 200
    assert body == {"status": "duplicate"}
    assert len(seen) == 1, "the retried activity was handed to the session a second time"


def test_a_delivered_reply_is_reported_without_echoing_the_message_id():
    router = ReplyRouter()
    router.bind(identity(), ReplyRoute(chat_id=CHAT_ID))

    class Outcome:
        def to_result(self):
            return {"delivered": True, "body_length": 11, "message_id": "graph-message-1"}

    app = application(reply=lambda decision: Outcome())
    code, _, body, raw = call(app)
    assert code == 200
    assert body["reply_delivered"] is True
    assert body["reply_body_length"] == 11
    assert b"graph-message-1" not in raw, "a Graph message id was echoed to the caller"


def test_a_lifecycle_activity_is_accepted_even_though_there_is_nothing_to_answer():
    """No reply is correct here, and it is a 200: the request itself succeeded."""
    app = application(
        reply=lambda decision: plan_reply_for(
            decision=decision, router=routed(), render=lambda activity: "hi"
        )
    )
    code, _, body, _ = call(app, body=payload_bytes(activity_payload(activity_type="typing")))
    assert code == 200
    assert body["status"] == "not_delivered"
    assert body["reply_delivered"] is False


def test_a_failed_send_is_a_200_because_a_redelivery_could_not_help():
    """5xx would make Bot Framework retry, and the retry is suppressed as a duplicate."""
    sent = []

    def send(arguments):
        sent.append(arguments)
        raise InboundTeamsError("service_error")

    app = application(
        reply=lambda decision: deliver_reply(
            plan_reply_for(decision=decision, router=routed(), render=lambda activity: "answer"),
            send=send,
        )
    )
    first, _, body, _ = call(app)
    assert first == 200
    assert body["reply_delivered"] is False
    assert body["reply_error"]
    assert len(sent) == 1

    # The retry that a 5xx would have provoked: suppressed, and no second send attempted.
    second, _, retry_body, _ = call(app)
    assert second == 200
    assert retry_body == {"status": "duplicate"}
    assert len(sent) == 1, "the retry triggered a second send"


def test_a_reply_failure_does_not_leak_the_chat_identifier():
    app = application(
        reply=lambda decision: (_ for _ in ()).throw(RuntimeError("chat 19:secret@thread.v2 boom"))
    )
    code, _, body, raw = call(app)
    assert code == 200
    assert b"19:secret@thread.v2" not in raw
    assert body["reply_delivered"] is False


def test_no_reply_callable_means_the_activity_is_only_acknowledged():
    code, _, body, _ = call(application(reply=None))
    assert code == 200
    assert body == {"status": "accepted"}


# ------------------------------------------------------------------ configuration


def test_an_empty_audience_is_refused_at_build_time():
    with pytest.raises(ValueError):
        build_application(registry=registry_with_binding(), audience="")


def test_a_relative_path_is_refused_at_build_time():
    with pytest.raises(ValueError):
        build_application(
            registry=registry_with_binding(), audience=AUDIENCE, path="teams/activities"
        )


# ------------------------------------------------------------------ over a real socket


class _QuietHandler(WSGIRequestHandler):
    """Silence the per-request log line; the assertions are the output."""

    def log_message(self, format, *args):  # noqa: A002 - mirrors the base signature
        pass


@pytest.mark.parametrize(
    "request_headers,expected_code",
    [
        ({"Authorization": BEARER, "Content-Type": "application/json"}, 200),
        ({"Content-Type": "application/json"}, 401),
        ({"Authorization": BEARER, "Content-Type": "text/plain"}, 415),
    ],
)
def test_the_endpoint_answers_over_a_real_http_socket(request_headers, expected_code):
    """Not a WSGI call: a real server on a real port, driven by a real HTTP client."""
    import http.client

    app = application()
    server = make_server("127.0.0.1", 0, app, handler_class=_QuietHandler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        connection = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
        body = payload_bytes()
        headers = {**request_headers, "Content-Length": str(len(body))}
        connection.request("POST", DEFAULT_WEBHOOK_PATH, body=body, headers=headers)
        response = connection.getresponse()
        raw = response.read()
        assert response.status == expected_code
        assert json.loads(raw.decode("utf-8"))["status"]
        connection.close()
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=10)


def test_the_endpoint_answers_a_duplicate_over_a_real_socket():
    import http.client

    app = application()
    server = make_server("127.0.0.1", 0, app, handler_class=_QuietHandler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        body = payload_bytes()
        statuses = []
        for _ in range(2):
            connection = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
            connection.request(
                "POST",
                DEFAULT_WEBHOOK_PATH,
                body=body,
                headers={
                    "Authorization": BEARER,
                    "Content-Type": "application/json",
                    "Content-Length": str(len(body)),
                },
            )
            response = connection.getresponse()
            statuses.append(response.status)
            response.read()
            connection.close()
        assert statuses == [200, 200]
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=10)


def test_a_get_to_the_real_endpoint_is_refused():
    import http.client

    app = application()
    server = make_server("127.0.0.1", 0, app, handler_class=_QuietHandler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        connection = http.client.HTTPConnection("127.0.0.1", port, timeout=10)
        connection.request("GET", DEFAULT_WEBHOOK_PATH)
        response = connection.getresponse()
        assert response.status == 405
        assert response.getheader("Allow") == "POST"
        response.read()
        connection.close()
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=10)
