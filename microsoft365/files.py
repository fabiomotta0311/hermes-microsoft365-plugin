"""Bounded, exact drive ``/content`` transfer (WP4).

What this module is
-------------------
The semantics layer over the verified ``/content`` contracts in
:mod:`microsoft365.sdk_contract`. It adds the four things the contracts deliberately do not
own: **validated addressing**, **exact bytes in both directions**, **bounds applied before
allocation**, and one **explicit, actionable status** for what the product does not implement.

Addressing
----------
A drive item is addressed exactly the way the generated builder is documented to accept it:
``"<root|item-id>:/<relative/path>"``. :func:`drive_item_address` builds that literal from a
validated path (or passes a validated opaque item id through untouched), and it is the *only*
join this module performs -- the URL is always rendered by the generated builder, which also
percent-encodes reserved and Unicode characters. Nothing here ever mangles a path by hand or
reads/writes a local file: the bytes travel to and from Graph and nowhere else.

A relative path is refused when it is absolute (``/a.txt``), a drive-absolute or already
addressed form (``root:/a.txt``, ``C:/a/b.txt``, a full URL), carries an empty segment
(``a//b``, ``a/``), a ``.``/``..`` traversal segment, a backslash, a colon, or any control
character. A path that only contains reserved or non-ASCII characters (a space, ``#``,
``?``, accents) is accepted verbatim and encoded by the SDK.

Bounds (10 MiB simple-transfer contract)
----------------------------------------
:data:`SIMPLE_TRANSFER_LIMIT_BYTES` is the product's simple-transfer bound, applied in **both**
directions, and :func:`encoded_length_bound` translates it into the base64 length that can
possibly hold it. Upload checks that encoded length **before** decoding (so an oversized
payload is never decoded, let alone allocated), then re-checks the decoded size -- the encoded
bound is a bound, not a proof: one extra byte encodes to the same length. Download checks the
item's *declared* size from authenticated metadata before issuing the byte request, and
re-checks the received length, so an unbounded response is never materialized silently.

Anything above the bound is reported as **not implemented** instead of being silently
attempted or truncated: an oversized upload carries ``UPLOAD_SESSION_NOT_IMPLEMENTED`` (the
drive upload session is a separate, unimplemented contract) and an oversized download carries
``DOWNLOAD_RANGE_NOT_IMPLEMENTED``. Both are :class:`TransferError` payloads with a static
message from the canonical taxonomy, an actionable ``action`` string, and bounded
``limit_bytes``/``size_bytes`` metadata.

Metadata, never caller input
----------------------------
The download's content type is read from the item's authenticated metadata (the verified
drives item read) and from nowhere else: :func:`download_file` has no ``content_type``
parameter at all, and it refuses metadata that is not a ``DriveItem``, that declares no usable
size, or that is not a file (no mime type -- a folder, for instance).

Execution seam
--------------
Both functions default to :func:`microsoft365.execution.execute_request` and take a
``client`` whose adapter is the client's own injected ``request_adapter``. ``execute`` exists
so the transfer logic is observable offline with a recording seam; production never overrides
it, and tests pin that default. Each request is built twice -- once through the verified
contract for the pre-execution assertion, once by the seam's generated caller -- which mirrors
the deliberate double build documented in :mod:`microsoft365.execution`; neither build performs
I/O and both are the same builder, configuration and body.
"""
from __future__ import annotations

import base64
import binascii
import unicodedata
from typing import Any, Callable
from urllib.parse import urlsplit

from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection
from kiota_abstractions.request_adapter import RequestAdapter

from . import execution
from .errors import MESSAGES, GraphError
from .results import BinaryResult, binary_result
from .sdk_contract import build_content_request_information, build_item_request_information

#: The product's simple-transfer contract: 10 MiB, in both directions. A caller may only
#: *lower* it; asking for more fails closed instead of silently widening the contract.
SIMPLE_TRANSFER_LIMIT_BYTES = 10 * 1024 * 1024

#: Default identifier of a path-addressed drive item (``root`` = the drive's root).
DEFAULT_PATH_IDENTIFIER = "root"

MAX_PATH_LENGTH = 4000
MAX_IDENTIFIER_LENGTH = 256
MAX_DRIVE_ID_LENGTH = 400
MAX_CONTENT_TYPE_LENGTH = 200

#: Explicit, actionable transfer statuses (never a silent attempt or a silent truncation).
UPLOAD_SESSION_NOT_IMPLEMENTED = "upload_session_not_implemented"
DOWNLOAD_RANGE_NOT_IMPLEMENTED = "download_range_not_implemented"
DECLARED_SIZE_MISMATCH = "declared_size_mismatch"
NOT_A_FILE = "not_a_file"

INVALID_TRANSFER_LIMIT = "invalid_transfer_limit"
INVALID_DRIVE_ID = "invalid_drive_id"
INVALID_DRIVE_PATH = "invalid_drive_path"
INVALID_ITEM_ID = "invalid_item_id"
INVALID_PATH_IDENTIFIER = "invalid_path_identifier"
ADDRESSING_REQUIRED = "addressing_required"
ADDRESSING_CONFLICT = "addressing_conflict"
INVALID_REQUEST_ADAPTER = "invalid_request_adapter"
METADATA_NOT_A_DRIVE_ITEM = "metadata_not_a_drive_item"
METADATA_WITHOUT_SIZE = "metadata_without_size"
METADATA_TARGET_MISMATCH = "metadata_target_mismatch"
CONTENT_TARGET_MISMATCH = "content_target_mismatch"
CONTENT_RESPONSE_NOT_BYTES = "content_response_not_bytes"
INVALID_CONTENT_BASE64 = "invalid_content_base64"
INVALID_CONTENT_TYPE = "invalid_content_type"

#: Static (never interpolated) next step per status: what a caller can actually do about it.
_ACTIONS = {
    UPLOAD_SESSION_NOT_IMPLEMENTED: "upload the file through a drive upload session instead",
    DOWNLOAD_RANGE_NOT_IMPLEMENTED: "download the file in ranges instead of one content request",
    DECLARED_SIZE_MISMATCH: "re-read the item metadata and retry the download",
    NOT_A_FILE: "address a file, not a folder or a non-file drive item",
}

TransferSeam = Callable[..., Any]


class TransferError(GraphError):
    """A sanitized transfer failure carrying an explicit, machine-readable ``status``.

    ``category`` comes from the canonical taxonomy and ``message`` is that taxonomy's static
    template, so no raw SDK/Graph text, path, file name or payload fragment can reach a
    caller. ``status`` says precisely which transfer rule was hit, and the optional bounded
    integers/action make the failure actionable without interpolating anything.
    """

    def __init__(
        self,
        category: str,
        *,
        status: str,
        action: str | None = None,
        limit_bytes: int | None = None,
        size_bytes: int | None = None,
        encoded_length: int | None = None,
    ):
        super().__init__(category, MESSAGES[category])
        self.status = str(status)
        self.action = action if action is not None else _ACTIONS.get(self.status)
        self.limit_bytes = limit_bytes
        self.size_bytes = size_bytes
        self.encoded_length = encoded_length

    def to_result(self) -> dict:
        payload = super().to_result()
        payload["status"] = self.status
        if self.action:
            payload["action"] = self.action
        if self.limit_bytes is not None:
            payload["limit_bytes"] = self.limit_bytes
        if self.size_bytes is not None:
            payload["size_bytes"] = self.size_bytes
        if self.encoded_length is not None:
            payload["encoded_length"] = self.encoded_length
        return payload


def _reject(status: str, **metadata: Any) -> TransferError:
    return TransferError("validation_error", status=status, **metadata)


def _has_control_characters(text: str) -> bool:
    """True for any Unicode control/format/private/unassigned character (category ``C*``)."""
    return any(unicodedata.category(character).startswith("C") for character in text)


def _require_text(value: Any, *, status: str, maximum: int) -> str:
    """A bounded, non-blank string: never stripped in place, never coerced."""
    if not isinstance(value, str) or value != value.strip() or not value:
        raise _reject(status)
    if len(value) > maximum:
        raise _reject(status)
    return value


def validate_drive_id(drive_id: Any) -> str:
    """A drive id is a single opaque token: no whitespace, no separators, no control chars."""
    text = _require_text(drive_id, status=INVALID_DRIVE_ID, maximum=MAX_DRIVE_ID_LENGTH)
    if _has_control_characters(text) or "/" in text or ":" in text:
        raise _reject(INVALID_DRIVE_ID)
    if any(character.isspace() for character in text):
        raise _reject(INVALID_DRIVE_ID)
    return text


def validate_drive_identifier(identifier: Any) -> str:
    """The ``<root|item-id>`` prefix of a path address."""
    text = _require_text(identifier, status=INVALID_PATH_IDENTIFIER, maximum=MAX_IDENTIFIER_LENGTH)
    if _has_control_characters(text) or "/" in text or ":" in text:
        raise _reject(INVALID_PATH_IDENTIFIER)
    if any(character.isspace() for character in text):
        raise _reject(INVALID_PATH_IDENTIFIER)
    return text


def validate_drive_item_id(drive_item_id: Any) -> str:
    """A plain, opaque drive item id (a path address is validated by :func:`validate_item_address`)."""
    text = _require_text(drive_item_id, status=INVALID_ITEM_ID, maximum=MAX_IDENTIFIER_LENGTH)
    if _has_control_characters(text) or "/" in text or ":" in text:
        raise _reject(INVALID_ITEM_ID)
    if any(character.isspace() for character in text):
        raise _reject(INVALID_ITEM_ID)
    return text


def validate_relative_drive_path(path: Any) -> str:
    """Validate a relative, forward-slash-separated path, and return it unchanged.

    Nothing is escaped, stripped or normalized here: reserved and Unicode characters are the
    generated builder's business. Every rejection is explicit so a caller can tell an absolute
    path from a traversal, an empty segment, a drive-absolute form or a control character.
    """
    text = _require_text(path, status=INVALID_DRIVE_PATH, maximum=MAX_PATH_LENGTH)
    if "\\" in text:
        raise _reject(INVALID_DRIVE_PATH)
    if text.startswith("/"):
        raise _reject(INVALID_DRIVE_PATH)
    if text.endswith("/") or "//" in text:
        raise _reject(INVALID_DRIVE_PATH)
    if ":" in text:
        # An addressing prefix, a drive letter or a URL is a drive-absolute form, never a
        # relative path, and would collide with the address delimiters below.
        raise _reject(INVALID_DRIVE_PATH)
    for segment in text.split("/"):
        if segment in {".", ".."}:
            raise _reject(INVALID_DRIVE_PATH)
    if _has_control_characters(text):
        raise _reject(INVALID_DRIVE_PATH)
    return text


def validate_item_address(address: Any) -> str:
    """Validate a prebuilt ``"<root|item-id>:/<relative/path>"`` address and return it as-is."""
    if not isinstance(address, str) or not address.endswith(":") or address != address.strip():
        raise _reject(INVALID_DRIVE_PATH)
    prefix, separator, relative_path = address[:-1].partition(":/")
    if not separator or not relative_path:
        raise _reject(INVALID_DRIVE_PATH)
    validate_drive_identifier(prefix)
    validate_relative_drive_path(relative_path)
    return address


def drive_item_address(
    *,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: Any = DEFAULT_PATH_IDENTIFIER,
) -> str:
    """Build (or validate) the drive item addressing string the generated builder takes.

    Exactly one of ``drive_item_id`` and ``path`` is required. With ``path``, the documented
    ``"<root|item-id>:/<relative/path>"`` literal is produced from the validated parts -- this
    is addressing, not URL building: the generated builder quotes the whole segment. With
    ``drive_item_id``, an already addressed value (ending in ``:``) is validated as an address
    and a plain id is validated as a plain id.
    """
    if drive_item_id is not None and path is not None:
        raise _reject(ADDRESSING_CONFLICT)
    if drive_item_id is not None:
        if isinstance(drive_item_id, str) and drive_item_id.endswith(":"):
            return validate_item_address(drive_item_id)
        return validate_drive_item_id(drive_item_id)
    if path is not None:
        return f"{validate_drive_identifier(identifier)}:/{validate_relative_drive_path(path)}:"
    raise _reject(ADDRESSING_REQUIRED)


def encoded_length_bound(limit_bytes: Any) -> int:
    """The longest base64 text that can possibly carry ``limit_bytes`` bytes.

    A bound, not a proof: ``limit_bytes`` and ``limit_bytes + 1`` share the same encoded
    length when the limit is not a multiple of three, which is exactly why the decoded size is
    checked again after decoding.
    """
    return 4 * ((_validated_limit(limit_bytes) + 2) // 3)


def _validated_limit(limit_bytes: Any) -> int:
    if isinstance(limit_bytes, bool) or not isinstance(limit_bytes, int):
        raise TransferError("configuration_error", status=INVALID_TRANSFER_LIMIT)
    if not 1 <= limit_bytes <= SIMPLE_TRANSFER_LIMIT_BYTES:
        raise TransferError(
            "configuration_error",
            status=INVALID_TRANSFER_LIMIT,
            limit_bytes=SIMPLE_TRANSFER_LIMIT_BYTES,
        )
    return limit_bytes


def validate_content_type(content_type: Any) -> str:
    """An explicit, plain media type: no CR/LF (header injection), no spaces, type/subtype."""
    text = _require_text(
        content_type, status=INVALID_CONTENT_TYPE, maximum=MAX_CONTENT_TYPE_LENGTH
    )
    if _has_control_characters(text) or any(character.isspace() for character in text):
        raise TransferError("validation_error", status=INVALID_CONTENT_TYPE)
    kind, separator, subtype = text.partition("/")
    if not separator or not kind or not subtype or "/" in subtype:
        raise TransferError("validation_error", status=INVALID_CONTENT_TYPE)
    return text


def _decode_base64_text(encoded: Any, *, limit_bytes: int) -> bytes:
    """Decode base64 under the encoded-length gate, then under the decoded-size gate.

    The order is the contract: an oversized payload is refused *before* decoding (and before
    the payload is even validated as base64), so nothing above the bound is ever decoded or
    allocated, and the failure says the simple transfer is not implemented rather than
    blaming the caller's encoding.

    ``validate=True`` is what makes the payload *base64* rather than something that merely
    needs base64's permissive alphabet filtering: without it ``"YWJj$"`` silently decodes to
    ``b"abc"`` and ``"  "`` to an empty file, so an upload could carry bytes no caller ever
    encoded.
    """
    if not isinstance(encoded, str):
        raise TransferError("validation_error", status=INVALID_CONTENT_BASE64)
    if len(encoded) > encoded_length_bound(limit_bytes):
        raise TransferError(
            "operation_not_implemented",
            status=UPLOAD_SESSION_NOT_IMPLEMENTED,
            limit_bytes=limit_bytes,
            encoded_length=len(encoded),
        )
    try:
        payload = base64.b64decode(encoded, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise TransferError("validation_error", status=INVALID_CONTENT_BASE64) from exc
    if len(payload) > limit_bytes:
        raise TransferError(
            "operation_not_implemented",
            status=UPLOAD_SESSION_NOT_IMPLEMENTED,
            limit_bytes=limit_bytes,
            size_bytes=len(payload),
        )
    return payload


def _request_path(request: Any) -> str:
    return urlsplit(str(getattr(request, "url", "") or "")).path


def _require_content_request(request: Any) -> Any:
    """Fail closed unless the verified contract rendered a request targeting the file bytes."""
    if not _request_path(request).endswith("/content"):
        raise TransferError("configuration_error", status=CONTENT_TARGET_MISMATCH)
    return request


def _require_metadata_request(request: Any) -> Any:
    """Fail closed unless the verified contract rendered an item *metadata* request."""
    if _request_path(request).endswith("/content"):
        raise TransferError("configuration_error", status=METADATA_TARGET_MISMATCH)
    return request


def _request_adapter(client: Any) -> RequestAdapter:
    adapter = getattr(client, "request_adapter", None)
    if not isinstance(adapter, RequestAdapter):
        raise TransferError("configuration_error", status=INVALID_REQUEST_ADAPTER)
    return adapter


def _empty_configuration() -> RequestConfiguration:
    # A real ``HeadersCollection``: a plain dict raises inside ``configure()`` (invariant 13).
    return RequestConfiguration(headers=HeadersCollection())


def _content_configuration(content_type: str) -> RequestConfiguration:
    headers = HeadersCollection()
    headers.add("Content-Type", content_type)
    return RequestConfiguration(headers=headers)


def _item_builder(client: Any, drive_id: str, address: str):
    return client.drives.by_drive_id(drive_id).items.by_drive_item_id(address)


def _content_builder(client: Any, drive_id: str, address: str):
    return _item_builder(client, drive_id, address).content


def _declared_metadata(item: Any) -> tuple[int, str, str | None]:
    """Read the download's contract from authenticated item metadata only.

    A response that is not a ``DriveItem``, one that declares no usable size, and one with no
    file facet/mime type (a folder) are all refused: none of them can bound or type a byte
    transfer, and a caller-supplied content type is never an alternative.
    """
    from msgraph.generated.models.drive_item import DriveItem

    if not isinstance(item, DriveItem):
        raise TransferError("configuration_error", status=METADATA_NOT_A_DRIVE_ITEM)
    size = item.size
    if isinstance(size, bool) or not isinstance(size, int) or size < 0:
        raise TransferError("configuration_error", status=METADATA_WITHOUT_SIZE)
    mime_type = getattr(item.file, "mime_type", None) if item.file is not None else None
    if not isinstance(mime_type, str) or not mime_type.strip() or _has_control_characters(mime_type):
        raise TransferError("validation_error", status=NOT_A_FILE)
    name = item.name if isinstance(item.name, str) and item.name.strip() else None
    return size, mime_type.strip(), name


def _response_header(response: Any, name: str) -> Any:
    headers = getattr(response, "headers", None)
    if headers is None:
        return None
    if hasattr(headers, "get"):
        value = headers.get(name)
        if value is not None:
            return value
        value = headers.get(name.lower())
        if value is not None:
            return value
    try:
        for key, value in headers.items():
            if str(key).lower() == name.lower():
                return value
    except AttributeError:
        pass
    return None


def _stream_response(response: Any, *, limit_bytes: int) -> bytes:
    """Read a response incrementally, stopping at the first byte over the bound."""
    declared = _response_header(response, "Content-Length")
    if declared is not None:
        try:
            declared = int(declared)
        except (TypeError, ValueError):
            declared = None
        if declared is not None and declared > limit_bytes:
            raise TransferError(
                "operation_not_implemented",
                status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
                limit_bytes=limit_bytes,
                size_bytes=declared,
            )

    if isinstance(response, (bytes, bytearray)):
        if len(response) > limit_bytes:
            raise TransferError(
                "operation_not_implemented",
                status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
                limit_bytes=limit_bytes,
                size_bytes=len(response),
            )
        return bytes(response)

    iterator = getattr(response, "iter_bytes", None)
    if not callable(iterator):
        iterator = getattr(response, "iter_content", None)
    chunks = iterator() if callable(iterator) else None
    data = bytearray()
    if chunks is not None:
        for chunk in chunks:
            if not isinstance(chunk, (bytes, bytearray)):
                raise TransferError("configuration_error", status=CONTENT_RESPONSE_NOT_BYTES)
            data.extend(chunk)
            if len(data) > limit_bytes:
                raise TransferError(
                    "operation_not_implemented",
                    status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
                    limit_bytes=limit_bytes,
                    size_bytes=limit_bytes + 1,
                )
        return bytes(data)

    reader = getattr(response, "read", None)
    if callable(reader):
        while len(data) <= limit_bytes:
            chunk = reader(limit_bytes - len(data) + 1)
            if not chunk:
                return bytes(data)
            if not isinstance(chunk, (bytes, bytearray)):
                raise TransferError("configuration_error", status=CONTENT_RESPONSE_NOT_BYTES)
            data.extend(chunk)
            if len(data) > limit_bytes:
                raise TransferError(
                    "operation_not_implemented",
                    status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
                    limit_bytes=limit_bytes,
                    size_bytes=limit_bytes + 1,
                )

    raise TransferError("configuration_error", status=CONTENT_RESPONSE_NOT_BYTES)


def download_file(
    client: Any,
    *,
    drive_id: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: Any = DEFAULT_PATH_IDENTIFIER,
    execute: TransferSeam = execution.execute_request,
    limit_bytes: Any = SIMPLE_TRANSFER_LIMIT_BYTES,
) -> BinaryResult:
    """Download one drive file's bytes and return them, base64-preserved, with a bound.

    The content type, the declared size and the name come from the item's authenticated
    metadata; the byte request is the verified ``/content`` GET for the same address. A
    declared size above the bound fails before the byte request is issued, and the received
    length is checked again against the bound and against the declared size.
    """
    drive = validate_drive_id(drive_id)
    address = drive_item_address(drive_item_id=drive_item_id, path=path, identifier=identifier)
    bound = _validated_limit(limit_bytes)
    adapter = _request_adapter(client)

    # No I/O before every argument and both request artifacts are validated.
    _require_metadata_request(
        build_item_request_information(
            client, "drive_item_read", drive_id=drive, drive_item_id=address
        )
    )
    _require_content_request(
        build_content_request_information(
            client, "download_files", drive_id=drive, drive_item_id=address
        )
    )

    item = execute(
        _item_builder(client, drive, address),
        method="GET",
        configuration=_empty_configuration(),
        adapter=adapter,
    )
    declared_size, mime_type, name = _declared_metadata(item)
    if declared_size > bound:
        raise TransferError(
            "operation_not_implemented",
            status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
            limit_bytes=bound,
            size_bytes=declared_size,
        )

    payload = execute(_content_builder(client, drive, address), method="GET", adapter=adapter)
    data = _stream_response(payload, limit_bytes=bound)
    if len(data) != declared_size:
        raise TransferError(
            "precondition_failed", status=DECLARED_SIZE_MISMATCH, size_bytes=len(data)
        )
    return binary_result(data, content_type=mime_type, name=name)


def upload_file(
    client: Any,
    *,
    drive_id: Any,
    content_base64: Any,
    content_type: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: Any = DEFAULT_PATH_IDENTIFIER,
    execute: TransferSeam = execution.execute_request,
    limit_bytes: Any = SIMPLE_TRANSFER_LIMIT_BYTES,
) -> dict:
    """Upload base64 content as the exact bytes of one drive file's ``/content``.

    The payload is bounded before it is decoded: an encoded length that cannot fit the bound
    is reported as ``upload_session_not_implemented`` without decoding anything, and the
    decoded length is checked again afterwards. The PUT carries the exact bytes and an
    explicit, validated ``Content-Type``.
    """
    drive = validate_drive_id(drive_id)
    address = drive_item_address(drive_item_id=drive_item_id, path=path, identifier=identifier)
    bound = _validated_limit(limit_bytes)
    media_type = validate_content_type(content_type)
    payload = _decode_base64_text(content_base64, limit_bytes=bound)
    adapter = _request_adapter(client)

    _require_content_request(
        build_content_request_information(
            client,
            "upload_files",
            drive_id=drive,
            drive_item_id=address,
            content=payload,
            content_type=media_type,
        )
    )

    execute(
        _content_builder(client, drive, address),
        method="PUT",
        configuration=_content_configuration(media_type),
        body=payload,
        adapter=adapter,
    )
    return {"status": "uploaded", "size_bytes": len(payload), "content_type": media_type}
