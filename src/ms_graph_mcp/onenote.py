"""Graph OneNote tools — notebooks, section groups, sections, pages, resources, page creation.

The read tools are enough to walk a whole OneNote library: notebooks contain
sections and section groups, section groups nest, sections hold pages, and a
page's HTML points at its images. Every list takes an opt-in ``page_token``
cursor, and page content an ``offset``, so nothing stops at a page size or a
character cap. Without those arguments each tool answers exactly as it always
has.
"""

from __future__ import annotations

import base64
import re
import urllib.parse

import httpx2
from pydantic import BaseModel, Field

from ms_graph_mcp.client import (
    GraphPayloadTooLarge,
    graph_get,
    graph_get_bytes,
    graph_get_text,
    graph_get_url,
    graph_post_raw,
)
from ms_graph_mcp.errors import graph_error_response, invalid_arguments, too_large
from ms_graph_mcp.odata import validate_graph_id
from ms_graph_mcp.tooling import READ_ONLY, WRITE_CREATE, tool

_PAGE_TOKEN_DESCRIPTION = (
    'Opt-in paging: "" for the first page, then the next_page_token returned (same ids). '
    "Omit for the plain list."
)

# One page of content: what notes_get_page_content returns per call.
_CONTENT_CHUNK = 50_000

# Largest embedded resource notes_get_resource returns. Base64 inflates it by a
# third, which keeps one result inside an ordinary MCP message.
_RESOURCE_MAX_BYTES = 5 * 1024 * 1024

# The part of a Graph nextLink before the OneNote collection path. Graph mints
# these and may spell the signed-in user as "me", "users/{id}" or
# "users('{upn}')", so all three are accepted. What the check pins is that the
# link stays inside OneNote and on the collection being paged; whose OneNote is
# decided by the delegated token, as for every other notes_ call.
_ONENOTE_LINK_PREFIX = re.compile(r"^/v1\.0/(?:me|users/[^/]+|users\('[^/']+'\))/onenote/")


class GetNotebooksInput(BaseModel):
    max_results: int = Field(10, description="Maximum notebooks to return (per page)")
    page_token: str | None = Field(None, description=_PAGE_TOKEN_DESCRIPTION)


class GetSectionsInput(BaseModel):
    notebook_id: str | None = Field(
        None,
        description="Notebook whose own sections to list. Exactly one of the two ids.",
    )
    section_group_id: str | None = Field(
        None,
        description="Section group whose sections to list. Exactly one of the two ids.",
    )
    page_token: str | None = Field(None, description=_PAGE_TOKEN_DESCRIPTION)


class GetSectionGroupsInput(BaseModel):
    notebook_id: str | None = Field(
        None,
        description="Notebook whose top-level section groups to list. Exactly one of the two ids.",
    )
    section_group_id: str | None = Field(
        None,
        description="Section group whose nested groups to list. Exactly one of the two ids.",
    )
    page_token: str | None = Field(None, description=_PAGE_TOKEN_DESCRIPTION)


class SaveToOnenoteInput(BaseModel):
    notebook_id: str = Field(description="Target notebook ID")
    section_id: str = Field(description="Target section ID")
    page_title: str = Field(description="Title of the new OneNote page")
    content_html: str = Field(description="HTML content for the page body")


def _parent(params: GetSectionsInput | GetSectionGroupsInput) -> tuple[str, str] | dict:
    """``("notebooks" | "sectionGroups", validated id)``, or an error to return.

    Exactly one parent: sections and section groups both hang off either a
    notebook or a section group, and guessing which one was meant would list
    the wrong thing without saying so.
    """
    given = [
        (collection, value, name)
        for collection, value, name in (
            ("notebooks", params.notebook_id, "notebook_id"),
            ("sectionGroups", params.section_group_id, "section_group_id"),
        )
        if value
    ]
    if len(given) != 1:
        return invalid_arguments(
            "Pass exactly one of notebook_id and section_group_id "
            f"(got {' and '.join(g[2] for g in given) or 'neither'})."
        )
    collection, value, name = given[0]
    try:
        return collection, validate_graph_id(value, name)
    except ValueError as exc:
        return invalid_arguments(str(exc))


def _check_page_token(page_token: str, collection: str) -> dict | None:
    """Refuse a cursor that is not a Graph nextLink for exactly this collection.

    ``graph_get_url`` already pins the host to Graph. This narrows it further,
    to the OneNote collection the caller is paging: a token is a URL the caller
    hands back, so without the check a notes_ tool could be steered into
    reading mail or files with the same delegated token.
    """
    parts = urllib.parse.urlsplit(page_token)
    path = urllib.parse.unquote(parts.path)
    prefix = _ONENOTE_LINK_PREFIX.match(path)
    if (
        parts.scheme != "https"
        or parts.netloc != "graph.microsoft.com"
        or parts.fragment
        or ".." in path
        or prefix is None
        or path[prefix.end() :].lower() != collection.lower()
    ):
        return invalid_arguments(
            "page_token is not a continuation of this list. Pass the next_page_token from the "
            "previous call unchanged, with the same ids, or an empty string to start over."
        )
    return None


async def _list(
    token: str,
    collection: str,
    query: dict,
    page_token: str | None,
    *,
    tool_name: str,
) -> dict:
    """Fetch one Graph page of a OneNote collection → raw JSON, or an error dict.

    ``collection`` is the path under ``/me/onenote/``. A non-empty
    ``page_token`` is followed instead of building the first request, after
    :func:`_check_page_token` has confirmed it belongs to ``collection``.
    """
    try:
        if page_token:
            refused = _check_page_token(page_token, collection)
            if refused:
                return refused
            return await graph_get_url(token, page_token)
        return await graph_get(token, f"/me/onenote/{collection}", **query)
    except httpx2.HTTPStatusError as exc:
        return graph_error_response(exc, scope="Notes.Read", tool=tool_name)


def _shape(data: dict, row, page_token: str | None) -> list[dict] | dict:
    """The old plain list, or ``{items, next_page_token}`` when paging was asked for."""
    if "error" in data:
        return data
    items = [row(item) for item in (data.get("value") or [])]
    if page_token is None:
        return items
    return {"items": items, "next_page_token": data.get("@odata.nextLink") or None}


def _web_url(obj: dict) -> str:
    return (obj.get("links") or {}).get("oneNoteWebUrl", {}).get("href", "")


def _named(obj: dict) -> dict:
    return {
        "id": obj.get("id", ""),
        "name": obj.get("displayName", ""),
        "last_modified": obj.get("lastModifiedDateTime", ""),
    }


@tool(
    description=(
        "List the signed-in user's OneNote notebooks, owned and shared, with id, name, "
        "last-modified date and web URL. Start here: a notebook id is what notes_list_sections "
        "and notes_list_section_groups take. Pass page_token (empty string first) to get "
        "{items, next_page_token} and reach every notebook; omit it for a plain list. "
        "Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=("get_notebooks",),
)
async def notes_list_notebooks(params: GetNotebooksInput, context: dict) -> list[dict] | dict:
    data = await _list(
        context["access_token"],
        "notebooks",
        {"$select": "id,displayName,lastModifiedDateTime,links", "$top": params.max_results},
        params.page_token,
        tool_name="notes_list_notebooks",
    )
    return _shape(data, lambda nb: {**_named(nb), "url": _web_url(nb)}, params.page_token)


@tool(
    description=(
        "List the sections in a OneNote notebook or section group (exactly one of notebook_id, "
        "section_group_id), with id, name and last-modified date. Pages live in sections, so a "
        "section id is what notes_list_pages and notes_create_page take. Sections inside "
        "section groups are reached via notes_list_section_groups. page_token (empty string "
        "first) returns {items, next_page_token}. Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=("get_sections",),
)
async def notes_list_sections(params: GetSectionsInput, context: dict) -> list[dict] | dict:
    parent = _parent(params)
    if isinstance(parent, dict):
        return parent
    kind, parent_id = parent
    data = await _list(
        context["access_token"],
        f"{kind}/{parent_id}/sections",
        {"$select": "id,displayName,lastModifiedDateTime"},
        params.page_token,
        tool_name="notes_list_sections",
    )
    return _shape(data, _named, params.page_token)


@tool(
    description=(
        "List the section groups (folders of sections) in a OneNote notebook or in another "
        "section group (exactly one of notebook_id, section_group_id), with id, name and "
        "last-modified date. Groups nest: recurse with this tool and notes_list_sections to "
        "reach every section a notebook holds. page_token (empty string first) returns "
        "{items, next_page_token}. Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=(),
)
async def notes_list_section_groups(
    params: GetSectionGroupsInput, context: dict
) -> list[dict] | dict:
    parent = _parent(params)
    if isinstance(parent, dict):
        return parent
    kind, parent_id = parent
    data = await _list(
        context["access_token"],
        f"{kind}/{parent_id}/sectionGroups",
        {"$select": "id,displayName,lastModifiedDateTime"},
        params.page_token,
        tool_name="notes_list_section_groups",
    )
    return _shape(data, _named, params.page_token)


async def create_onenote_page(
    token: str, *, section_id: str, page_title: str, content_html: str
) -> dict:
    """Create a OneNote page in a section and return the RAW Graph page object.

    The single OneNote page-create HTTP — callers (the ``save_to_onenote`` tool,
    the action-sink, REST routes) all go through here instead of hand-rolling the
    POST, so the content-type (text/html) and page wrapper stay consistent.
    """
    html_body = f"""<!DOCTYPE html>
<html>
<head><title>{page_title}</title></head>
<body>{content_html}</body>
</html>"""

    return await graph_post_raw(
        token,
        f"/me/onenote/sections/{section_id}/pages",
        html_body.encode("utf-8"),
        "text/html",
    )


@tool(
    description=(
        "Create a new page in a OneNote section and return its id and web URL. Takes a section id "
        "from notes_list_sections, a title, and the body as valid HTML — plain text will not "
        "render correctly. Use for saving meeting notes, summaries or research into the user's "
        "notebooks. Always adds a new page; it cannot edit an existing one. "
        "Requires Notes.Create."
    ),
    annotations=WRITE_CREATE,
    aliases=("save_to_onenote",),
)
async def notes_create_page(params: SaveToOnenoteInput, context: dict) -> dict:
    data = await create_onenote_page(
        context["access_token"],
        section_id=params.section_id,
        page_title=params.page_title,
        content_html=params.content_html,
    )
    return {
        "id": data.get("id", ""),
        "title": data.get("title", ""),
        "url": _web_url(data),
        "created_at": data.get("createdDateTime", ""),
    }


class ListOnenotePagesInput(BaseModel):
    section_id: str = Field(
        description="The section id to list pages from, from notes_list_sections"
    )
    max_results: int = Field(
        20, ge=1, le=100, description="Maximum pages to return per call (1-100)"
    )
    page_token: str | None = Field(None, description=_PAGE_TOKEN_DESCRIPTION)


class GetOnenotePageInput(BaseModel):
    page_id: str = Field(description="The page id, from notes_list_pages")
    offset: int = Field(
        0,
        ge=0,
        description="Character offset to read from: 0, then each next_offset returned.",
    )


class GetOnenoteResourceInput(BaseModel):
    resource_id: str = Field(
        description="Id between 'resources/' and '/$value' in a page's <img src> URL; not the URL"
    )


@tool(
    description=(
        "List the pages in one OneNote section, newest-modified first, with id, title, creation "
        "date and web URL. Takes a section id from notes_list_sections. Titles only — "
        "notes_get_page_content reads a page. Pass page_token (empty string first) to get "
        "{items, next_page_token} and walk every page, up to 100 per call; omit it for a plain "
        "list. Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=(),
)
async def notes_list_pages(params: ListOnenotePagesInput, context: dict) -> list[dict] | dict:
    try:
        section_id = validate_graph_id(params.section_id, "section_id")
    except ValueError as exc:
        return invalid_arguments(str(exc))
    data = await _list(
        context["access_token"],
        f"sections/{section_id}/pages",
        {
            "$select": "id,title,createdDateTime,lastModifiedDateTime,links",
            "$top": params.max_results,
            "$orderby": "lastModifiedDateTime desc",
        },
        params.page_token,
        tool_name="notes_list_pages",
    )
    return _shape(
        data,
        lambda p: {
            "id": p.get("id", ""),
            "title": p.get("title", ""),
            "created_at": p.get("createdDateTime", ""),
            "last_modified": p.get("lastModifiedDateTime", ""),
            "url": _web_url(p),
        },
        params.page_token,
    )


@tool(
    description=(
        "Read one OneNote page's body as HTML, how OneNote stores it, given a page id from "
        "notes_list_pages. Returns up to 50,000 characters from offset; while next_offset is not "
        "null, call again with offset=next_offset to read the rest. Images are <img> tags whose "
        "bytes notes_get_resource fetches. Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=(),
)
async def notes_get_page_content(params: GetOnenotePageInput, context: dict) -> dict:
    token = context["access_token"]
    try:
        page_id = validate_graph_id(params.page_id, "page_id")
    except ValueError as exc:
        return invalid_arguments(str(exc))
    # Page content is HTML, not JSON — graph_get_text is the same helper the
    # meeting-transcript VTT path uses.
    try:
        html = await graph_get_text(token, f"/me/onenote/pages/{page_id}/content")
    except httpx2.HTTPStatusError as exc:
        return graph_error_response(exc, scope="Notes.Read", tool="notes_get_page_content")
    # Graph has no range read for page content, so each chunk re-fetches the page
    # and slices it. ``truncated`` keeps its old meaning: more follows this chunk.
    end = params.offset + _CONTENT_CHUNK
    more = len(html) > end
    return {
        "page_id": page_id,
        "content_html": html[params.offset : end],
        "truncated": more,
        "length": len(html),
        "offset": params.offset,
        "next_offset": end if more else None,
    }


@tool(
    description=(
        "Download one image or file embedded in a OneNote page, as base64 with its MIME type. "
        "Takes the resource id: the path segment after resources/ in an <img src> or <object "
        "data> URL in notes_get_page_content's HTML, never the URL. Refuses anything over 5 MB. "
        "Page text comes from notes_get_page_content. Requires Notes.Read."
    ),
    annotations=READ_ONLY,
    aliases=(),
)
async def notes_get_resource(params: GetOnenoteResourceInput, context: dict) -> dict:
    try:
        resource_id = validate_graph_id(params.resource_id, "resource_id")
    except ValueError as exc:
        return invalid_arguments(str(exc))
    if "/" in resource_id:
        # The id regex admits "/" for base64 segments elsewhere; here it would
        # let "{id}/$value" style input address another path under resources/.
        return invalid_arguments(
            "resource_id is one path segment — the part between 'resources/' and '/$value' — "
            "not a URL or a path."
        )
    try:
        raw, content_type = await graph_get_bytes(
            context["access_token"],
            f"/me/onenote/resources/{resource_id}/$value",
            max_bytes=_RESOURCE_MAX_BYTES,
        )
    except GraphPayloadTooLarge as exc:
        return too_large("The OneNote resource", size=exc.size, limit=exc.limit)
    except httpx2.HTTPStatusError as exc:
        return graph_error_response(exc, scope="Notes.Read", tool="notes_get_resource")
    return {
        "resource_id": resource_id,
        "mime_type": content_type.split(";")[0].strip() or "application/octet-stream",
        "size": len(raw),
        "content_b64": base64.b64encode(raw).decode("ascii"),
    }
