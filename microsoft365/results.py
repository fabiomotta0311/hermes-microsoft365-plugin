"""Sanitized Graph model and binary result normalization."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Mapping

from .errors import GraphError, is_sensitive_key, is_structural_key, to_graph_error

_MODEL_FIELDS = (
    "id", "name", "display_name", "subject", "title", "content", "size", "web_url",
    "created_date_time", "last_modified_date_time", "status", "plan_id", "bucket_id",
)


@dataclass(frozen=True)
class BinaryResult:
    content_base64: str
    content_type: str
    size: int
    name: str | None = None


def normalize_collection_page(
    values: Any,
    *,
    odata_count: Any = None,
    next_link: Any = None,
    max_items: int = 100,
    max_string: int = 4000,
) -> dict:
    """Normalize a graph collection page (or several pages) into one stable envelope.

    Pagination needs the collection-level fields -- ``value``, ``@odata.count`` and
    ``@odata.nextLink`` -- *and* every per-item guarantee :func:`normalize_result` already
    provides (secret redaction, bounded strings, binary artifacts outside string
    truncation). Items are therefore normalized one by one with that same function instead
    of being copied into a second, weaker normalizer.

    Only a link the caller has already validated belongs in ``next_link``: this function
    reports what it is given and does not decide whether a continuation is safe. The
    envelope is display/storage data -- the authoritative continuation is the paginator's
    own report, because a very long link is truncated here by ``max_string``.
    """
    items = [] if values is None else list(values)
    envelope: dict = {
        "value": [
            normalize_result(item, max_items=max_items, max_string=max_string)
            for item in items[:max_items]
        ]
    }
    if odata_count is not None:
        envelope["@odata.count"] = normalize_result(odata_count, max_items=max_items, max_string=max_string)
    if next_link:
        envelope["@odata.nextLink"] = normalize_result(next_link, max_items=max_items, max_string=max_string)
    return envelope


def normalize_result(
    value: Any,
    *,
    max_items: int = 100,
    max_string: int = 4000,
    max_depth: int = 8,
    _depth: int = 0,
    _seen: set[int] | None = None,
) -> Any:
    if isinstance(value, BinaryResult):
        # Base64 is an integrity-bearing transport field, not display text.
        return {key: item for key, item in asdict(value).items() if item is not None}
    if isinstance(value, GraphError):
        return value.to_result()
    if isinstance(value, BaseException):
        # A raw SDK/Graph/Azure exception is never rendered as text: it goes through the
        # canonical taxonomy, which never copies its message, headers, body or request.
        return to_graph_error(value).to_result()
    if _seen is None:
        _seen = set()
    if _depth > max_depth:
        return "[TRUNCATED]"
    if value is None or type(value) in {bool, int, float}:
        return value
    if isinstance(value, str):
        return value if len(value) <= max_string else value[:max_string] + "...[TRUNCATED]"
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {"type": "bytes", "size": len(value)}

    identity = id(value)
    if identity in _seen:
        return "[CYCLE]"
    _seen.add(identity)
    try:
        recurse = lambda item: normalize_result(
            item,
            max_items=max_items,
            max_string=max_string,
            max_depth=max_depth,
            _depth=_depth + 1,
            _seen=_seen,
        )
        if isinstance(value, Mapping):
            output = {}
            for key, item in list(value.items())[:max_items]:
                # Redaction rules live in the canonical taxonomy module, so the error path
                # and the result path can never drift apart.
                output[str(key)] = (
                    "[REDACTED]" if is_sensitive_key(key) or is_structural_key(key) else recurse(item)
                )
            return output
        if isinstance(value, (list, tuple, set)):
            return [recurse(item) for item in list(value)[:max_items]]

        output = {}
        if hasattr(value, "value"):
            output["value"] = recurse(getattr(value, "value") or [])
        if getattr(value, "odata_count", None) is not None:
            output["@odata.count"] = recurse(value.odata_count)
        if getattr(value, "odata_next_link", None):
            output["@odata.nextLink"] = recurse(value.odata_next_link)
        for field_name in _MODEL_FIELDS:
            if hasattr(value, field_name):
                item = getattr(value, field_name)
                if item is not None:
                    output[field_name] = recurse(item)
        if hasattr(value, "additional_data"):
            output["additional_data"] = recurse(getattr(value, "additional_data") or {})
        return output if output else recurse(str(value))
    finally:
        _seen.discard(identity)
