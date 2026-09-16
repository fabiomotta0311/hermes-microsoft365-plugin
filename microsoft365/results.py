"""Sanitized Graph model and binary result normalization."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Mapping

_SECRET_PARTS = ("secret", "token", "authorization", "cookie", "password", "api_key", "apikey")
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


def _secret_key(key: Any) -> bool:
    normalized = str(key).lower().replace("-", "_")
    return any(part in normalized for part in _SECRET_PARTS)


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
                output[str(key)] = "[REDACTED]" if _secret_key(key) or str(key).lower() in {
                    "headers", "request_information", "response"
                } else recurse(item)
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
