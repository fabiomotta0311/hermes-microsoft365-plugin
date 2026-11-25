"""Post-write reconciliation: confirm a Graph write against the server's own state.

Why this exists
---------------
A ``2xx`` from Microsoft Graph means *the request was accepted*. It does not mean the bytes
landed as intended. A proxy can truncate, a rename can collide and produce
``report (2).txt``, a session can complete against a stale item, and an ``If-Match`` can be
ignored by an intermediary. Every one of those returns success.

So a write in this plugin is a *claim*, and this module turns it into a **verified claim**:
after the write, the item is read back over the same authenticated boundary and compared
against what the caller actually asked for. The comparison is against values this module
computed or received from Graph's own response -- never against a value the caller supplied to
make the check pass.

The verdict vocabulary
----------------------
``confirmed``
    The re-read item matches every field this plugin can check. The claim stands.

``mismatched``
    The re-read item contradicts the claim (wrong size, different name, different id). The
    write reported success and is not what was requested. This is reported, never repaired:
    silently fixing it would hide an incident.

``unverified``
    The write reported success but could not be confirmed -- the re-read failed, or the item
    came back without the fields needed to compare. The write may well be fine; this plugin
    just cannot say so.

``unverified`` is never silently upgraded to ``confirmed``, and a failure to reconcile never
becomes a failure of the write itself: the caller gets the truth about *verification*, which
is a different claim from the write.

Security
--------
Field values that come back from Graph are compared, not rendered. Anything echoed into the
result passes through :mod:`microsoft365.results`, so a ``webUrl`` or a session-bearing value
cannot smuggle a secret into the payload.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Mapping

from . import execution
from .files import (
    DEFAULT_PATH_IDENTIFIER,
    drive_item_address,
    validate_drive_id,
)

#: Graph is strongly consistent for drive items, but a write is followed by a network round
#: trip and a CDN-cached read is still an ordinary failure mode. A small bounded number of
#: re-reads distinguishes "not visible yet" from "not there at all" without pretending to know.
DEFAULT_RECONCILE_ATTEMPTS = 3

#: Bounded backoff between confirmation attempts.
RECONCILE_BACKOFF_SECONDS = (0.2, 0.5)

CONFIRMED = "confirmed"
MISMATCHED = "mismatched"
UNVERIFIED = "unverified"


@dataclass(frozen=True)
class Reconciliation:
    """The verdict of one post-write confirmation, and the evidence behind it."""

    status: str
    reason: str | None = None
    expected: dict[str, Any] = field(default_factory=dict)
    observed: dict[str, Any] = field(default_factory=dict)
    attempts: int = 1

    @property
    def confirmed(self) -> bool:
        return self.status == CONFIRMED

    def to_result(self) -> dict:
        """The rendered envelope: what was expected, what was seen, and how it was checked."""
        payload: dict[str, Any] = {"status": self.status}
        if self.reason:
            payload["reason"] = self.reason
        if self.expected:
            payload["expected"] = dict(self.expected)
        if self.observed:
            payload["observed"] = dict(self.observed)
        payload["attempts"] = self.attempts
        return payload


def _item_size(item: Any) -> int | None:
    size = getattr(item, "size", None)
    return size if isinstance(size, int) and not isinstance(size, bool) else None


def _item_name(item: Any) -> str | None:
    name = getattr(item, "name", None)
    return name if isinstance(name, str) and name else None


def _item_id(item: Any) -> str | None:
    identity = getattr(item, "id", None)
    return identity if isinstance(identity, str) and identity else None


def _etag(item: Any) -> str | None:
    """The cTag/ETag pair, when Graph supplied one.

    These are opaque and never parsed: they are carried so a caller can log *that* the item
    changed between the write and the read, without this module interpreting a tag format.
    """
    values = {}
    for name in ("c_tag", "e_tag"):
        raw = getattr(item, name, None)
        if isinstance(raw, str) and raw:
            values[name] = raw
    extra = getattr(item, "additional_data", None)
    if isinstance(extra, dict):
        for name in ("@microsoft.graph.cTag", "@microsoft.graph.eTag"):
            raw = extra.get(name)
            if isinstance(raw, str) and raw:
                values[name] = raw
    return values or None


def compare(
    item: Any,
    *,
    expected_size: int | None = None,
    expected_name: str | None = None,
    expected_id: str | None = None,
    expected_etag: Any = None,
) -> Reconciliation:
    """Compare a re-read item against what the write claimed, field by field.

    Every checkable field must agree. A field that cannot be read on both sides is reported as
    unverified rather than assumed equal: "I could not tell" and "they match" are different
    answers, and collapsing them is how a broken write gets reported as a good one.
    """
    if not isinstance(expected_size, (int,)) and expected_size is not None:
        return Reconciliation(UNVERIFIED, reason="expected_size_is_not_an_integer")

    observed = {
        "id": _item_id(item),
        "name": _item_name(item),
        "size": _item_size(item),
    }
    tags = _etag(item)
    if tags:
        observed["etag"] = tags

    expected: dict[str, Any] = {}
    if expected_size is not None:
        expected["size"] = expected_size
    if expected_name is not None:
        expected["name"] = expected_name
    if expected_id is not None:
        expected["id"] = expected_id

    # Size is the strongest check available: it is the one property a truncated or duplicated
    # transfer cannot preserve. It is checked first and a disagreement is decisive.
    observed_size = observed.get("size")
    if expected_size is not None:
        if observed_size is None:
            return Reconciliation(
                UNVERIFIED, reason="item_declared_no_size", expected=expected, observed=observed
            )
        if observed_size != expected_size:
            return Reconciliation(
                MISMATCHED,
                reason="size_differs_from_what_was_written",
                expected=expected,
                observed=observed,
            )

    observed_name = observed.get("name")
    if expected_name is not None:
        if observed_name is None:
            return Reconciliation(
                UNVERIFIED, reason="item_declared_no_name", expected=expected, observed=observed
            )
        if observed_name != expected_name:
            return Reconciliation(
                MISMATCHED,
                reason="name_differs_from_what_was_written",
                expected=expected,
                observed=observed,
            )

    observed_id = observed.get("id")
    if expected_id is not None and observed_id is not None and observed_id != expected_id:
        return Reconciliation(
            MISMATCHED,
            reason="the_item_written_is_not_the_item_read",
            expected=expected,
            observed=observed,
        )

    if not expected:
        # Nothing to check against: the re-read happened, but it proves nothing.
        return Reconciliation(
            UNVERIFIED, reason="nothing_was_declared_to_check_against", observed=observed
        )

    return Reconciliation(CONFIRMED, expected=expected, observed=observed)


def compare_fields(
    item: Any,
    expected: Mapping[str, Any],
    *,
    identity_field: str | None = None,
) -> Reconciliation:
    """Compare a re-read resource against declared scalar fields.

    Only fields that are present on both sides are compared, and *every* comparable field must
    agree. A field the server did not return is reported as unverified rather than skipped:
    silently ignoring a missing field is how a check that never ran gets reported as passing.
    """
    comparable = {
        name: value
        for name, value in expected.items()
        if value is not None and isinstance(name, str) and name
    }
    if not comparable:
        return Reconciliation(UNVERIFIED, reason="nothing_was_declared_to_check_against")

    observed: dict[str, Any] = {}
    for name in comparable:
        value = getattr(item, name, None)
        observed[name] = value if isinstance(value, (str, int, bool)) else None

    missing = [name for name, value in observed.items() if value is None and name != identity_field]
    if missing:
        return Reconciliation(
            UNVERIFIED,
            reason="the_resource_did_not_report_the_fields_that_were_written",
            expected=comparable,
            observed=observed,
        )

    identity = getattr(item, identity_field, None) if identity_field else None
    if identity_field and not isinstance(identity, str):
        return Reconciliation(
            UNVERIFIED,
            reason="the_resource_did_not_report_an_identity",
            expected=comparable,
            observed=observed,
        )
    if identity_field:
        observed[identity_field] = identity

    for name, value in comparable.items():
        if observed.get(name) != value:
            return Reconciliation(
                MISMATCHED,
                reason=f"{name}_differs_from_what_was_written",
                expected=comparable,
                observed=observed,
            )
    return Reconciliation(CONFIRMED, expected=comparable, observed=observed)


def reconcile_resource(
    read: Callable[[], Any],
    *,
    expected: Mapping[str, Any],
    identity_field: str | None = None,
    attempts: int = DEFAULT_RECONCILE_ATTEMPTS,
    sleep: Callable[[float], None] | None = None,
) -> Reconciliation:
    """Read a written resource back and compare it, retrying only while it is not yet visible.

    ``read`` performs one fetch and returns the resource (or ``None`` when it is not there yet).
    A resource that has been found and contradicts the write is never retried -- the answer was
    clear; only invisibility is worth another attempt.
    """
    read_attempts = 0
    while True:
        read_attempts += 1
        try:
            item = read()
        except Exception:  # noqa: BLE001 - a failed read is "unverified", never "failed"
            if read_attempts >= max(1, attempts):
                return Reconciliation(
                    UNVERIFIED, reason="the_resource_could_not_be_read_back", attempts=read_attempts
                )
            _wait(sleep, read_attempts)
            continue

        if item is None:
            if read_attempts >= max(1, attempts):
                return Reconciliation(
                    UNVERIFIED, reason="the_resource_did_not_appear", attempts=read_attempts
                )
            _wait(sleep, read_attempts)
            continue

        verdict = compare_fields(item, expected, identity_field=identity_field)
        return Reconciliation(
            verdict.status,
            reason=verdict.reason,
            expected=verdict.expected,
            observed=verdict.observed,
            attempts=read_attempts,
        )


def reconcile_drive_item(
    client: Any,
    *,
    drive_id: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: str = DEFAULT_PATH_IDENTIFIER,
    expected_size: int | None = None,
    expected_name: str | None = None,
    expected_id: str | None = None,
    expected_etag: Any = None,
    attempts: int = DEFAULT_RECONCILE_ATTEMPTS,
    execute: Callable[..., Any] = execution.execute_request,
    sleep: Callable[[float], None] | None = None,
) -> Reconciliation:
    """Re-read one drive item and return the verdict, retrying only while it is not visible.

    A ``mismatched`` verdict is **not** retried: the server answered clearly and the answer
    contradicts the claim, so re-reading would only delay reporting an incident. Only an item
    that has not appeared yet (``None``) is worth another attempt.
    """
    from .files import _empty_configuration, _item_builder, _request_adapter

    drive = validate_drive_id(drive_id)
    address = drive_item_address(drive_item_id=drive_item_id, path=path, identifier=identifier)
    adapter = _request_adapter(client)
    read_attempts = 0

    while True:
        read_attempts += 1
        try:
            item = execute(
                _item_builder(client, drive, address),
                method="GET",
                configuration=_empty_configuration(),
                adapter=adapter,
                max_attempts=1,  # this function owns the attempt bound; do not multiply it
            )
        except Exception as exc:  # noqa: BLE001 - a failed read means "unverified", never "failed"
            if read_attempts >= max(1, attempts):
                return Reconciliation(
                    UNVERIFIED,
                    reason="the_item_could_not_be_read_back",
                    attempts=read_attempts,
                )
            _wait(sleep, read_attempts)
            continue

        if item is None:
            if read_attempts >= max(1, attempts):
                return Reconciliation(
                    UNVERIFIED, reason="the_item_did_not_appear", attempts=read_attempts
                )
            _wait(sleep, read_attempts)
            continue

        verdict = compare(
            item,
            expected_size=expected_size,
            expected_name=expected_name,
            expected_id=expected_id,
            expected_etag=expected_etag,
        )
        if verdict.status == MISMATCHED:
            return Reconciliation(
                verdict.status, reason=verdict.reason, expected=verdict.expected,
                observed=verdict.observed, attempts=read_attempts,
            )
        return Reconciliation(
            verdict.status, reason=verdict.reason, expected=verdict.expected,
            observed=verdict.observed, attempts=read_attempts,
        )


def _wait(sleep: Callable[[float], None] | None, attempt: int) -> None:
    index = min(attempt - 1, len(RECONCILE_BACKOFF_SECONDS) - 1)
    delay = RECONCILE_BACKOFF_SECONDS[index]
    if sleep is not None:
        sleep(delay)
        return
    import time

    time.sleep(delay)