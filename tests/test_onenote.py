"""OneNote read tools: walking a whole library without losing anything.

A client importing someone's OneNote has to reach sections inside section
groups, page past Graph's page size, read pages longer than one result, and
fetch the images a page points at. Each of those is opt-in, so the cases that
matter most are the ones a reviewer misses: today's arguments still get today's
shape, the paging cursor cannot be steered outside the collection it came from,
and every id that reaches a Graph path is validated.
"""

from __future__ import annotations

import base64
from unittest.mock import AsyncMock, patch

import httpx2
import pytest
from pydantic import ValidationError

from ms_graph_mcp.client import GraphPayloadTooLarge, graph_get_bytes
from ms_graph_mcp.onenote import (
    GetNotebooksInput,
    GetOnenotePageInput,
    GetOnenoteResourceInput,
    GetSectionGroupsInput,
    GetSectionsInput,
    ListOnenotePagesInput,
    notes_get_page_content,
    notes_get_resource,
    notes_list_notebooks,
    notes_list_pages,
    notes_list_section_groups,
    notes_list_sections,
)

_CTX = {"access_token": "tok"}
_BASE = "https://graph.microsoft.com/v1.0"
_REAL_ASYNC_CLIENT = httpx2.AsyncClient


def _http_error(status: int) -> httpx2.HTTPStatusError:
    request = httpx2.Request("GET", f"{_BASE}/x")
    return httpx2.HTTPStatusError(
        "boom", request=request, response=httpx2.Response(status, request=request)
    )


def _named(i: int) -> dict:
    return {"id": f"0-ABC!{i}", "displayName": f"N{i}", "lastModifiedDateTime": "2026-01-01"}


# ── today's arguments, today's shape ──────────────────────────────────────────


async def test_notebooks_without_page_token_return_the_plain_list():
    nb = {**_named(1), "links": {"oneNoteWebUrl": {"href": "https://x/nb"}}}
    page = {"value": [nb], "@odata.nextLink": f"{_BASE}/me/onenote/notebooks?$skip=10"}
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value=page)) as get:
        result = await notes_list_notebooks(GetNotebooksInput(), _CTX)
    assert get.call_args.args[1] == "/me/onenote/notebooks"
    assert get.call_args.kwargs["$top"] == 10
    assert result == [
        {"id": "0-ABC!1", "name": "N1", "last_modified": "2026-01-01", "url": "https://x/nb"}
    ]


async def test_sections_by_notebook_id_keep_the_plain_list():
    with patch(
        "ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": [_named(2)]})
    ) as get:
        result = await notes_list_sections(GetSectionsInput(notebook_id="0-ABC!1"), _CTX)
    assert get.call_args.args[1] == "/me/onenote/notebooks/0-ABC!1/sections"
    assert result == [{"id": "0-ABC!2", "name": "N2", "last_modified": "2026-01-01"}]


async def test_pages_without_page_token_return_the_plain_list():
    p = {"id": "p1", "title": "T", "createdDateTime": "c", "lastModifiedDateTime": "m"}
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": [p]})) as get:
        result = await notes_list_pages(ListOnenotePagesInput(section_id="s1"), _CTX)
    assert get.call_args.args[1] == "/me/onenote/sections/s1/pages"
    assert get.call_args.kwargs["$top"] == 20
    assert result == [
        {"id": "p1", "title": "T", "created_at": "c", "last_modified": "m", "url": ""}
    ]


async def test_short_page_content_reads_whole_with_no_next_offset():
    with patch("ms_graph_mcp.onenote.graph_get_text", new=AsyncMock(return_value="<p>hi</p>")):
        result = await notes_get_page_content(GetOnenotePageInput(page_id="p1"), _CTX)
    assert result == {
        "page_id": "p1",
        "content_html": "<p>hi</p>",
        "truncated": False,
        "length": 9,
        "offset": 0,
        "next_offset": None,
    }


# ── section groups ────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "kwargs,path",
    [
        ({"notebook_id": "nb1"}, "/me/onenote/notebooks/nb1/sectionGroups"),
        ({"section_group_id": "sg1"}, "/me/onenote/sectionGroups/sg1/sectionGroups"),
    ],
)
async def test_section_groups_list_under_a_notebook_or_a_group(kwargs, path):
    with patch(
        "ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": [_named(3)]})
    ) as get:
        result = await notes_list_section_groups(GetSectionGroupsInput(**kwargs), _CTX)
    assert get.call_args.args[1] == path
    assert result == [{"id": "0-ABC!3", "name": "N3", "last_modified": "2026-01-01"}]


async def test_sections_inside_a_section_group():
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": []})) as get:
        await notes_list_sections(GetSectionsInput(section_group_id="sg1"), _CTX)
    assert get.call_args.args[1] == "/me/onenote/sectionGroups/sg1/sections"


@pytest.mark.parametrize(
    "fn,model",
    [(notes_list_sections, GetSectionsInput), (notes_list_section_groups, GetSectionGroupsInput)],
)
@pytest.mark.parametrize(
    "kwargs", [{}, {"notebook_id": "nb1", "section_group_id": "sg1"}, {"notebook_id": ""}]
)
async def test_exactly_one_parent_is_required(fn, model, kwargs):
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock()) as get:
        result = await fn(model(**kwargs), _CTX)
    get.assert_not_called()
    assert result["error"] == "INVALID_ARGUMENTS"
    assert "exactly one" in result["message"]


# ── ids are validated before they reach a path ────────────────────────────────


@pytest.mark.parametrize(
    "call",
    [
        lambda: notes_list_sections(GetSectionsInput(notebook_id="../me/messages"), _CTX),
        lambda: notes_list_sections(GetSectionsInput(section_group_id="a?b"), _CTX),
        lambda: notes_list_section_groups(GetSectionGroupsInput(notebook_id="a#b"), _CTX),
        lambda: notes_list_section_groups(GetSectionGroupsInput(section_group_id="a b"), _CTX),
        lambda: notes_list_pages(ListOnenotePagesInput(section_id="x/../y"), _CTX),
        lambda: notes_get_page_content(GetOnenotePageInput(page_id="p?$expand=x"), _CTX),
        lambda: notes_get_resource(GetOnenoteResourceInput(resource_id="../pages"), _CTX),
    ],
)
async def test_bad_ids_are_refused_without_a_graph_call(call):
    with (
        patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock()) as get,
        patch("ms_graph_mcp.onenote.graph_get_text", new=AsyncMock()) as get_text,
        patch("ms_graph_mcp.onenote.graph_get_bytes", new=AsyncMock()) as get_bytes,
    ):
        result = await call()
    for mock in (get, get_text, get_bytes):
        mock.assert_not_called()
    assert result["error"] == "INVALID_ARGUMENTS"


async def test_a_quoted_notebook_id_is_unquoted_like_the_other_ids():
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": []})) as get:
        await notes_list_sections(GetSectionsInput(notebook_id="'0-ABC!1'"), _CTX)
    assert get.call_args.args[1] == "/me/onenote/notebooks/0-ABC!1/sections"


async def test_a_graph_failure_comes_back_as_a_structured_error():
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(side_effect=_http_error(403))):
        result = await notes_list_notebooks(GetNotebooksInput(), _CTX)
    assert result["error"] == "SCOPE_DENIED"
    assert result["scope"] == "Notes.Read"


# ── paging ────────────────────────────────────────────────────────────────────


async def test_pages_round_trip_through_the_page_token():
    first_link = f"{_BASE}/me/onenote/sections/s1/pages?$top=2&$skip=2"
    first = {"value": [{"id": "p1"}, {"id": "p2"}], "@odata.nextLink": first_link}
    last = {"value": [{"id": "p3"}]}
    with (
        patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value=first)) as get,
        patch("ms_graph_mcp.onenote.graph_get_url", new=AsyncMock(return_value=last)) as get_url,
    ):
        one = await notes_list_pages(
            ListOnenotePagesInput(section_id="s1", max_results=2, page_token=""), _CTX
        )
        two = await notes_list_pages(
            ListOnenotePagesInput(section_id="s1", page_token=one["next_page_token"]), _CTX
        )
    assert get.call_args.kwargs["$top"] == 2
    assert [p["id"] for p in one["items"]] == ["p1", "p2"]
    assert one["next_page_token"] == first_link
    get_url.assert_awaited_once_with("tok", first_link)
    assert [p["id"] for p in two["items"]] == ["p3"]
    assert two["next_page_token"] is None


async def test_notebooks_page_and_end_with_a_null_token():
    with patch("ms_graph_mcp.onenote.graph_get", new=AsyncMock(return_value={"value": []})):
        result = await notes_list_notebooks(GetNotebooksInput(page_token=""), _CTX)
    assert result == {"items": [], "next_page_token": None}


@pytest.mark.parametrize(
    "link",
    [
        # Graph may spell the user out instead of "me", and percent-encode the "!".
        f"{_BASE}/users/1f2e-3d4c/onenote/sectionGroups/0-ABC%2119/sections?$skip=20",
        f"{_BASE}/users('alex@example.com')/onenote/sectionGroups/0-ABC!19/sections?$skip=20",
        f"{_BASE}/me/onenote/sectiongroups/0-ABC!19/sections?$skip=20",
    ],
)
async def test_graph_spellings_of_the_same_collection_are_accepted(link):
    with patch(
        "ms_graph_mcp.onenote.graph_get_url", new=AsyncMock(return_value={"value": []})
    ) as get_url:
        result = await notes_list_sections(
            GetSectionsInput(section_group_id="0-ABC!19", page_token=link), _CTX
        )
    get_url.assert_awaited_once_with("tok", link)
    assert result == {"items": [], "next_page_token": None}


@pytest.mark.parametrize(
    "token",
    [
        f"{_BASE}/me/messages?$skip=10",  # another workload, same token
        f"{_BASE}/me/onenote/sections/OTHER/pages?$skip=20",  # another section
        f"{_BASE}/me/onenote/sections/s1/pages/../../notebooks",
        f"{_BASE}/me/onenote/sections/s1/pages/p1/content",
        "https://graph.microsoft.com.evil.com/v1.0/me/onenote/sections/s1/pages",
        "http://graph.microsoft.com/v1.0/me/onenote/sections/s1/pages",
        f"{_BASE}/beta/me/onenote/sections/s1/pages",
        "/me/onenote/sections/s1/pages",
        "not a url",
    ],
)
async def test_a_page_token_for_anything_else_is_refused(token):
    with patch("ms_graph_mcp.onenote.graph_get_url", new=AsyncMock()) as get_url:
        result = await notes_list_pages(
            ListOnenotePagesInput(section_id="s1", page_token=token), _CTX
        )
    get_url.assert_not_called()
    assert result["error"] == "INVALID_ARGUMENTS"


def test_page_size_is_capped_at_graphs_maximum():
    assert ListOnenotePagesInput(section_id="s1", max_results=100).max_results == 100
    with pytest.raises(ValidationError):
        ListOnenotePagesInput(section_id="s1", max_results=101)


# ── long pages read in chunks ─────────────────────────────────────────────────


async def test_a_long_page_reads_back_whole_through_next_offset():
    html = "".join(f"<p>{i}</p>" for i in range(20_000))[:120_001]  # 50k + 50k + 20,001
    chunks, offset = [], 0
    with patch("ms_graph_mcp.onenote.graph_get_text", new=AsyncMock(return_value=html)):
        while offset is not None:
            result = await notes_get_page_content(
                GetOnenotePageInput(page_id="p1", offset=offset), _CTX
            )
            assert len(result["content_html"]) <= 50_000
            assert result["truncated"] is (result["next_offset"] is not None)
            assert result["length"] == len(html)
            chunks.append(result["content_html"])
            offset = result["next_offset"]
    assert len(chunks) == 3
    assert "".join(chunks) == html


async def test_an_offset_past_the_end_is_empty_and_final():
    with patch("ms_graph_mcp.onenote.graph_get_text", new=AsyncMock(return_value="<p>x</p>")):
        result = await notes_get_page_content(GetOnenotePageInput(page_id="p1", offset=1_000), _CTX)
    assert result["content_html"] == ""
    assert result["next_offset"] is None


def test_a_negative_offset_is_rejected():
    with pytest.raises(ValidationError):
        GetOnenotePageInput(page_id="p1", offset=-1)


# ── embedded resources ────────────────────────────────────────────────────────


async def test_a_resource_comes_back_base64_with_its_media_type():
    png = b"\x89PNG\r\n\x1a\nfake"
    with patch(
        "ms_graph_mcp.onenote.graph_get_bytes",
        new=AsyncMock(return_value=(png, "image/png; charset=binary")),
    ) as get:
        result = await notes_get_resource(GetOnenoteResourceInput(resource_id="0-abc!1-r"), _CTX)
    assert get.call_args.args[1] == "/me/onenote/resources/0-abc!1-r/$value"
    assert get.call_args.kwargs["max_bytes"] == 5 * 1024 * 1024
    assert result == {
        "resource_id": "0-abc!1-r",
        "mime_type": "image/png",
        "size": len(png),
        "content_b64": base64.b64encode(png).decode("ascii"),
    }


@pytest.mark.parametrize(
    "resource_id",
    [f"{_BASE}/me/onenote/resources/r1/$value", "r1/$value", "r1/content", "r1/../../pages"],
)
async def test_a_resource_url_or_path_is_not_an_id(resource_id):
    with patch("ms_graph_mcp.onenote.graph_get_bytes", new=AsyncMock()) as get:
        result = await notes_get_resource(GetOnenoteResourceInput(resource_id=resource_id), _CTX)
    get.assert_not_called()
    assert result["error"] == "INVALID_ARGUMENTS"


async def test_an_oversized_resource_is_refused_terminally():
    with patch(
        "ms_graph_mcp.onenote.graph_get_bytes",
        new=AsyncMock(side_effect=GraphPayloadTooLarge(9_000_000, 5 * 1024 * 1024)),
    ):
        result = await notes_get_resource(GetOnenoteResourceInput(resource_id="r1"), _CTX)
    assert result["error"] == "TOO_LARGE"
    assert result["retryable"] is False
    assert result["size"] == 9_000_000


async def test_a_missing_resource_is_not_found():
    with patch("ms_graph_mcp.onenote.graph_get_bytes", new=AsyncMock(side_effect=_http_error(404))):
        result = await notes_get_resource(GetOnenoteResourceInput(resource_id="r1"), _CTX)
    assert result["error"] == "NOT_FOUND"


# ── graph_get_bytes: the cap holds on the wire ────────────────────────────────


def _serve(handler):
    """Patch the client module's AsyncClient to answer from ``handler`` in-process."""

    def factory(**kwargs):
        return _REAL_ASYNC_CLIENT(transport=httpx2.MockTransport(handler), **kwargs)

    return patch("ms_graph_mcp.client.httpx2.AsyncClient", side_effect=factory)


async def test_get_bytes_returns_the_body_and_content_type():
    seen = {}

    def handler(request):
        seen["url"] = str(request.url)
        seen["auth"] = request.headers["Authorization"]
        return httpx2.Response(200, content=b"abc", headers={"Content-Type": "image/jpeg"})

    with _serve(handler):
        body, ctype = await graph_get_bytes("tok", "/me/onenote/resources/r1/$value", max_bytes=10)
    assert (body, ctype) == (b"abc", "image/jpeg")
    assert seen == {"url": f"{_BASE}/me/onenote/resources/r1/$value", "auth": "Bearer tok"}


async def test_get_bytes_refuses_a_declared_length_over_the_cap():
    with _serve(lambda request: httpx2.Response(200, content=b"x" * 11)):
        with pytest.raises(GraphPayloadTooLarge) as exc:
            await graph_get_bytes("tok", "/r", max_bytes=10)
    assert exc.value.size == 11


async def test_get_bytes_stops_an_undeclared_length_at_the_cap():
    async def body():
        for _ in range(5):
            yield b"x" * 4

    with _serve(lambda request: httpx2.Response(200, content=body())):
        with pytest.raises(GraphPayloadTooLarge) as exc:
            await graph_get_bytes("tok", "/r", max_bytes=10)
    assert exc.value.size is None


async def test_get_bytes_raises_on_an_error_status():
    with _serve(lambda request: httpx2.Response(404, json={"error": {"code": "x"}})):
        with pytest.raises(httpx2.HTTPStatusError):
            await graph_get_bytes("tok", "/r", max_bytes=10)
