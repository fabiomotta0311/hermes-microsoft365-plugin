"""SharePoint and OneDrive handlers: search, read, download_files and upload_files (WP7).

The same rules as the Outlook and Calendar modules apply: each request is built on the
generated SDK builders from the declared arguments and pinned against the verified
request-information contract of the operation (``microsoft365/sdk_contract.py``; see
``tests/test_handlers_files.py``). No URL is assembled by hand and no secret is read here.

Action-specific rules this module owns (what the shared argument contract cannot express)
---------------------------------------------------------------------------------------
* ``sharepoint.search`` scopes a search by one container at most. With no container it
  searches **sites** (``GET /sites?$search=…``); with ``drive_id`` it searches that drive
  (``GET /drives/{drive_id}/search(q='…')``); with ``site_id`` it resolves the site to its
  default document library first (``GET /sites/{site_id}/drive``) and then searches the
  resolved drive. ``onedrive.search`` is drive-scoped only: a ``site_id`` is refused (OneDrive
  is not site-scoped) and a search without ``drive_id`` is refused (there is no tenant-wide
  drive search in application mode).
* ``read`` addresses one drive item by an opaque ``item_id`` or a validated relative
  ``item_path``, through ``microsoft365.files.drive_item_address`` — the generated builder
  renders the path and percent-encodes it, never this module.
* ``download_files``/``upload_files`` go through the bounded ``microsoft365/files.py``
  transfer: exact bytes, the 10 MiB simple-transfer bound before allocation, the content type
  of a download read from authenticated metadata only, and an upload whose ``overwrite`` is
  ``False`` is refused because a simple ``PUT /content`` always replaces its target (a
  conflict-safe upload needs an upload session, which is not implemented).

The four reads above are executable; the two uploads are implemented, contract-pinned and
exercised offline through the seam, and deliberately **not executable** until the generic
host approval fix (CORE-1/CORE-2) lands (R5).
"""
from __future__ import annotations

from typing import Mapping

from ..execution import execute_request
from ..files import download_file, drive_item_address, upload_file
from ..paging import paginate
from . import (
    HandlerContext,
    collection_payload,
    handler,
    identifier,
    item_budget,
    open_graph_client,
    refuse,
    select,
    success_payload,
    typed_configuration,
)

SHAREPOINT_SEARCH = "sharepoint.search"
SHAREPOINT_READ = "sharepoint.read"
SHAREPOINT_DOWNLOAD = "sharepoint.download_files"
SHAREPOINT_UPLOAD = "sharepoint.upload_files"
ONEDRIVE_SEARCH = "onedrive.search"
ONEDRIVE_READ = "onedrive.read"
ONEDRIVE_DOWNLOAD = "onedrive.download_files"
ONEDRIVE_UPLOAD = "onedrive.upload_files"

#: The media type of an upload whose caller did not declare one. A simple ``PUT /content``
#: with a byte body has this as its natural, standard type.
DEFAULT_UPLOAD_CONTENT_TYPE = "application/octet-stream"

#: OData doubles a single quote inside a literal; the generated ``search_with_q`` path segment
#: then percent-encodes it. Matches ``microsoft365.sdk_contract`` (WP1).
_ODATA_ESCAPE = "'", "''"


def _required_query(arguments: Mapping, *, key: str) -> str:
    value = arguments.get("query")
    if not isinstance(value, str) or not value.strip():
        refuse("validation_error", f"{key}: query must be a non-blank search string")
    return value


def _drive_search(client, adapter, key: str, drive_id: str, query: str, limit: int) -> dict:
    escaped = query.replace(*_ODATA_ESCAPE)
    builder = client.drives.by_drive_id(drive_id).search_with_q(escaped)
    configuration = typed_configuration(builder, top=limit)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    return collection_payload(key, paged)


def _sites_search(client, adapter, key: str, query: str, limit: int) -> dict:
    builder = client.sites
    configuration = typed_configuration(builder, search=query, top=limit)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    return collection_payload(key, paged)


def _site_drive_search(client, adapter, key: str, site_id: str, query: str, limit: int) -> dict:
    drive = execute_request(client.sites.by_site_id(site_id).drive, method="GET", adapter=adapter)
    resolved = getattr(drive, "id", None)
    if not isinstance(resolved, str) or not resolved.strip():
        refuse("service_error", f"{key}: the site did not resolve to a drive with an id")
    return _drive_search(client, adapter, key, resolved, query, limit)


def _search(arguments: Mapping, context: HandlerContext, *, key: str, site_scoped: bool) -> dict:
    query = _required_query(arguments, key=key)
    drive_id = arguments.get("drive_id")
    site_id = arguments.get("site_id")
    if drive_id is not None and site_id is not None:
        # Re-checked here (the shared argument contract refuses it too): two containers would
        # mean two different searches and only one of them would be sent.
        refuse("validation_error", f"{key}: scope the search by drive_id or site_id, not both")
    limit = item_budget(arguments, operation=key)

    if drive_id is not None:
        resolved_drive = identifier(arguments, "drive_id", operation=key)
        client, adapter = open_graph_client(context)
        return _drive_search(client, adapter, key, resolved_drive, query, limit)
    if site_id is not None:
        if not site_scoped:
            refuse("validation_error", f"{key}: OneDrive is not site-scoped, so site_id is refused")
        resolved_site = identifier(arguments, "site_id", operation=key)
        client, adapter = open_graph_client(context)
        return _site_drive_search(client, adapter, key, resolved_site, query, limit)
    if not site_scoped:
        # OneDrive has no tenant-wide search: a drive is required to scope the search.
        refuse("validation_error", f"{key}: onedrive search requires drive_id to scope the search")
    client, adapter = open_graph_client(context)
    return _sites_search(client, adapter, key, query, limit)


@handler("sharepoint", "search")
def sharepoint_search(arguments: Mapping, context: HandlerContext) -> dict:
    """Search SharePoint sites, or one site's/drive's contents when scoped."""
    return _search(arguments, context, key=SHAREPOINT_SEARCH, site_scoped=True)


@handler("onedrive", "search")
def onedrive_search(arguments: Mapping, context: HandlerContext) -> dict:
    """Search one drive's contents, scoped by ``drive_id``."""
    return _search(arguments, context, key=ONEDRIVE_SEARCH, site_scoped=False)


# ---------------------------------------------------------------- reads


def _drive_read(arguments: Mapping, context: HandlerContext, *, key: str) -> dict:
    drive_id = identifier(arguments, "drive_id", operation=key)
    # Addressing is validated before the client exists (defence in depth); the transfer layer
    # re-validates it. Exactly one of item_id/item_path must be supplied.
    drive_item_address(drive_item_id=arguments.get("item_id"), path=arguments.get("item_path"))
    fields = select(arguments, operation=key)

    client, adapter = open_graph_client(context)
    builder = client.drives.by_drive_id(drive_id).items.by_drive_item_id(
        drive_item_address(drive_item_id=arguments.get("item_id"), path=arguments.get("item_path"))
    )
    configuration = typed_configuration(builder, select=fields)
    response = execute_request(builder, method="GET", configuration=configuration, adapter=adapter)
    return success_payload(key, response)


@handler("sharepoint", "read")
def sharepoint_read(arguments: Mapping, context: HandlerContext) -> dict:
    """Read one SharePoint drive item by id or by validated relative path."""
    return _drive_read(arguments, context, key=SHAREPOINT_READ)


@handler("onedrive", "read")
def onedrive_read(arguments: Mapping, context: HandlerContext) -> dict:
    """Read one OneDrive drive item by id or by validated relative path."""
    return _drive_read(arguments, context, key=ONEDRIVE_READ)


# ---------------------------------------------------------------- transfers


def _drive_download(arguments: Mapping, context: HandlerContext, *, key: str) -> dict:
    drive_id = identifier(arguments, "drive_id", operation=key)
    drive_item_address(drive_item_id=arguments.get("item_id"), path=arguments.get("item_path"))

    client, _adapter = open_graph_client(context)
    result = download_file(
        client,
        drive_id=drive_id,
        drive_item_id=arguments.get("item_id"),
        path=arguments.get("item_path"),
    )
    return success_payload(key, result)


@handler("sharepoint", "download_files")
def sharepoint_download_files(arguments: Mapping, context: HandlerContext) -> dict:
    """Download the exact bytes of one SharePoint file, bounded before allocation."""
    return _drive_download(arguments, context, key=SHAREPOINT_DOWNLOAD)


@handler("onedrive", "download_files")
def onedrive_download_files(arguments: Mapping, context: HandlerContext) -> dict:
    """Download the exact bytes of one OneDrive file, bounded before allocation."""
    return _drive_download(arguments, context, key=ONEDRIVE_DOWNLOAD)


def _drive_upload(arguments: Mapping, context: HandlerContext, *, key: str) -> dict:
    drive_id = identifier(arguments, "drive_id", operation=key)
    drive_item_address(drive_item_id=arguments.get("item_id"), path=arguments.get("item_path"))
    overwrite = arguments.get("overwrite")
    if overwrite is False:
        # A simple PUT /content always replaces its target; a conflict-safe upload needs an
        # upload session with a conflictBehavior, which this plugin does not implement.
        refuse(
            "validation_error",
            f"{key}: overwrite=False cannot be honored by a simple content upload, which always "
            "replaces the target; a conflict-safe upload needs an upload session, which is not "
            "implemented",
        )
    content_base64 = arguments.get("content_base64")
    content_type = arguments.get("content_type") or DEFAULT_UPLOAD_CONTENT_TYPE

    client, _adapter = open_graph_client(context)
    result = upload_file(
        client,
        drive_id=drive_id,
        content_base64=content_base64,
        content_type=content_type,
        drive_item_id=arguments.get("item_id"),
        path=arguments.get("item_path"),
    )
    return success_payload(key, result)


@handler("sharepoint", "upload_files")
def sharepoint_upload_files(arguments: Mapping, context: HandlerContext) -> dict:
    """Upload base64 content as the exact bytes of one SharePoint file (withheld)."""
    return _drive_upload(arguments, context, key=SHAREPOINT_UPLOAD)


@handler("onedrive", "upload_files")
def onedrive_upload_files(arguments: Mapping, context: HandlerContext) -> dict:
    """Upload base64 content as the exact bytes of one OneDrive file (withheld)."""
    return _drive_upload(arguments, context, key=ONEDRIVE_UPLOAD)
