"""Google Docs MCP server with SA delegation and tab support."""

import os
import json
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator

from mcp.server.fastmcp import FastMCP, Context
from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive.readonly",
]


def get_services():
    sa_path = os.environ.get("SERVICE_ACCOUNT_PATH")
    subject = os.environ.get("SUBJECT_EMAIL")
    if not sa_path:
        raise ValueError("SERVICE_ACCOUNT_PATH env var required")
    if not subject:
        raise ValueError("SUBJECT_EMAIL env var required")

    creds = service_account.Credentials.from_service_account_file(sa_path, scopes=SCOPES)
    creds = creds.with_subject(subject)

    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return docs_service, drive_service


@asynccontextmanager
async def lifespan(server: FastMCP) -> AsyncIterator[dict]:
    docs_service, drive_service = get_services()
    yield {"docs": docs_service, "drive": drive_service}


mcp = FastMCP("google-docs", lifespan=lifespan)


def _get_ctx(ctx: Context) -> tuple:
    lc = ctx.request_context.lifespan_context
    return lc["docs"], lc["drive"]


# --- Tab helpers ---

def _extract_text_from_body(body: dict) -> str:
    """Extract text from a doc body, preserving heading structure as markdown."""
    content = body.get("content", [])
    parts = []

    heading_map = {
        "HEADING_1": "# ",
        "HEADING_2": "## ",
        "HEADING_3": "### ",
        "HEADING_4": "#### ",
        "HEADING_5": "##### ",
        "HEADING_6": "###### ",
        "TITLE": "# ",
        "SUBTITLE": "## ",
    }

    for element in content:
        if "paragraph" not in element:
            continue
        para = element["paragraph"]
        style = para.get("paragraphStyle", {})
        named_style = style.get("namedStyleType", "NORMAL_TEXT")

        line_parts = []
        for elem in para.get("elements", []):
            text_run = elem.get("textRun")
            if text_run:
                line_parts.append(text_run.get("content", ""))

        line = "".join(line_parts)
        prefix = heading_map.get(named_style, "")
        line_stripped = line.rstrip("\n")
        if line_stripped:
            parts.append(prefix + line_stripped)
        elif not prefix:
            parts.append("")

    return "\n".join(parts)


def _body_end_index(body: dict) -> int:
    """Get the end index of a body (for appending)."""
    content = body.get("content", [])
    if not content:
        return 1
    last = content[-1]
    return last.get("endIndex", 1) - 1


def _flatten_tabs(tabs: list, depth: int = 0) -> list:
    """Flatten a nested tab tree into a list with depth info."""
    result = []
    for tab in tabs:
        props = tab.get("tabProperties", {})
        result.append({
            "tabId": props.get("tabId", ""),
            "title": props.get("title", ""),
            "index": props.get("index", 0),
            "depth": depth,
            "tab": tab,
        })
        for child in tab.get("childTabs", []):
            result.extend(_flatten_tabs([child], depth + 1))
    return result


def _find_tab(tabs: list, tab_id: str = "", tab_title: str = "") -> dict | None:
    """Find a tab by ID or title in the tab tree. Returns the tab dict or None."""
    flat = _flatten_tabs(tabs)
    for entry in flat:
        if tab_id and entry["tabId"] == tab_id:
            return entry["tab"]
        if tab_title and entry["title"].lower() == tab_title.lower():
            return entry["tab"]
    return None


def _get_doc_with_tabs(docs, document_id: str) -> dict:
    """Fetch a document with tab content included."""
    return docs.documents().get(
        documentId=document_id,
        includeTabsContent=True,
    ).execute()


def _resolve_tab(doc: dict, tab_id: str = "", tab_title: str = "") -> tuple[dict, str]:
    """Resolve a tab from the doc. Returns (tab_body, tab_id).
    If no tab specified, returns the first tab."""
    tabs = doc.get("tabs", [])
    if not tabs:
        # Fallback for docs without tabs
        return doc.get("body", {}), ""

    if tab_id or tab_title:
        tab = _find_tab(tabs, tab_id=tab_id, tab_title=tab_title)
        if not tab:
            raise ValueError(f"Tab not found: id={tab_id!r} title={tab_title!r}")
        props = tab.get("tabProperties", {})
        body = tab.get("documentTab", {}).get("body", {})
        return body, props.get("tabId", "")

    # Default: first tab
    first = tabs[0]
    props = first.get("tabProperties", {})
    body = first.get("documentTab", {}).get("body", {})
    return body, props.get("tabId", "")


# --- Tools ---

@mcp.tool()
def list_tabs(ctx: Context, document_id: str) -> str:
    """List all tabs in a Google Doc.

    Args:
        document_id: The document ID (from the URL)
    """
    docs, _ = _get_ctx(ctx)
    doc = _get_doc_with_tabs(docs, document_id)
    tabs = doc.get("tabs", [])
    flat = _flatten_tabs(tabs)

    tab_list = []
    for entry in flat:
        tab_list.append({
            "tabId": entry["tabId"],
            "title": entry["title"],
            "index": entry["index"],
            "depth": entry["depth"],
        })

    return json.dumps({
        "documentTitle": doc.get("title", "Untitled"),
        "tabs": tab_list,
        "count": len(tab_list),
    }, indent=2)


@mcp.tool()
def read_document(ctx: Context, document_id: str, tab_id: str = "", tab_title: str = "") -> str:
    """Read a Google Doc tab and return its content as markdown-formatted text.
    If no tab is specified, reads all tabs with headers.

    Args:
        document_id: The document ID (from the URL)
        tab_id: Optional tab ID to read a specific tab
        tab_title: Optional tab title to read a specific tab (case-insensitive)
    """
    docs, _ = _get_ctx(ctx)
    doc = _get_doc_with_tabs(docs, document_id)
    title = doc.get("title", "Untitled")
    tabs = doc.get("tabs", [])

    # If a specific tab requested, return just that tab
    if tab_id or tab_title:
        body, resolved_id = _resolve_tab(doc, tab_id=tab_id, tab_title=tab_title)
        tab = _find_tab(tabs, tab_id=resolved_id)
        tab_name = tab.get("tabProperties", {}).get("title", "") if tab else ""
        text = _extract_text_from_body(body)
        return f"# {title} — [{tab_name}]\n\n{text}"

    # Otherwise, read all tabs
    if not tabs:
        body = doc.get("body", {})
        text = _extract_text_from_body(body)
        return f"# {title}\n\n{text}"

    flat = _flatten_tabs(tabs)
    sections = [f"# {title}\n"]
    for entry in flat:
        tab_obj = entry["tab"]
        tab_name = entry["title"]
        body = tab_obj.get("documentTab", {}).get("body", {})
        text = _extract_text_from_body(body)
        heading_level = "#" * (entry["depth"] + 2)
        sections.append(f"{heading_level} [{tab_name}]\n\n{text}")

    return "\n\n---\n\n".join(sections)


@mcp.tool()
def get_document_info(ctx: Context, document_id: str) -> str:
    """Get metadata about a Google Doc (title, ID, link, tabs). Lightweight.

    Args:
        document_id: The document ID (from the URL)
    """
    docs, _ = _get_ctx(ctx)
    doc = _get_doc_with_tabs(docs, document_id)
    title = doc.get("title", "Untitled")
    doc_id = doc.get("documentId", document_id)
    link = f"https://docs.google.com/document/d/{doc_id}/edit"
    tabs = doc.get("tabs", [])
    flat = _flatten_tabs(tabs)
    tab_info = [{"tabId": e["tabId"], "title": e["title"]} for e in flat]

    return json.dumps({
        "title": title,
        "documentId": doc_id,
        "link": link,
        "tabs": tab_info,
    }, indent=2)


@mcp.tool()
def create_document(ctx: Context, title: str, body_text: str = "") -> str:
    """Create a new Google Doc.

    Args:
        title: Title of the new document
        body_text: Optional initial text content
    """
    docs, _ = _get_ctx(ctx)
    doc = docs.documents().create(body={"title": title}).execute()
    doc_id = doc.get("documentId")

    if body_text:
        docs.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [{"insertText": {"location": {"index": 1}, "text": body_text}}]},
        ).execute()

    link = f"https://docs.google.com/document/d/{doc_id}/edit"
    return json.dumps({"documentId": doc_id, "title": title, "link": link}, indent=2)


@mcp.tool()
def append_text(ctx: Context, document_id: str, text: str, tab_id: str = "", tab_title: str = "") -> str:
    """Append text to the end of a Google Doc tab.

    Args:
        document_id: The document ID
        text: Text to append
        tab_id: Optional tab ID to target (default: first tab)
        tab_title: Optional tab title to target (case-insensitive)
    """
    docs, _ = _get_ctx(ctx)
    doc = _get_doc_with_tabs(docs, document_id)
    body, resolved_tab_id = _resolve_tab(doc, tab_id=tab_id, tab_title=tab_title)
    end_idx = _body_end_index(body)

    if end_idx > 1:
        text = "\n" + text

    location = {"index": end_idx}
    if resolved_tab_id:
        location["tabId"] = resolved_tab_id

    docs.documents().batchUpdate(
        documentId=document_id,
        body={"requests": [{"insertText": {"location": location, "text": text}}]},
    ).execute()

    return json.dumps({"status": "ok", "appended_chars": len(text), "at_index": end_idx, "tabId": resolved_tab_id})


@mcp.tool()
def insert_text(ctx: Context, document_id: str, text: str, index: int, tab_id: str = "", tab_title: str = "") -> str:
    """Insert text at a specific character index in a Google Doc tab.

    Args:
        document_id: The document ID
        text: Text to insert
        index: Character index (1-based, 1 = start of doc)
        tab_id: Optional tab ID to target (default: first tab)
        tab_title: Optional tab title to target (case-insensitive)
    """
    docs, _ = _get_ctx(ctx)

    location = {"index": index}
    if tab_id:
        location["tabId"] = tab_id
    elif tab_title:
        doc = _get_doc_with_tabs(docs, document_id)
        _, resolved_tab_id = _resolve_tab(doc, tab_title=tab_title)
        if resolved_tab_id:
            location["tabId"] = resolved_tab_id

    docs.documents().batchUpdate(
        documentId=document_id,
        body={"requests": [{"insertText": {"location": location, "text": text}}]},
    ).execute()

    return json.dumps({"status": "ok", "inserted_chars": len(text), "at_index": index, "tabId": location.get("tabId", "")})


@mcp.tool()
def replace_text(ctx: Context, document_id: str, find: str, replace_with: str, match_case: bool = True, tab_id: str = "", tab_title: str = "") -> str:
    """Find and replace all occurrences of text in a Google Doc. Can target a specific tab.

    Args:
        document_id: The document ID
        find: Text to search for
        replace_with: Replacement text
        match_case: Whether to match case (default True)
        tab_id: Optional tab ID to limit replacement to
        tab_title: Optional tab title to limit replacement to (case-insensitive)
    """
    docs, _ = _get_ctx(ctx)

    request = {
        "replaceAllText": {
            "containsText": {"text": find, "matchCase": match_case},
            "replaceText": replace_with,
        }
    }

    # Resolve tab for scoping
    resolved_tab_id = tab_id
    if not resolved_tab_id and tab_title:
        doc = _get_doc_with_tabs(docs, document_id)
        _, resolved_tab_id = _resolve_tab(doc, tab_title=tab_title)

    if resolved_tab_id:
        request["replaceAllText"]["tabsCriteria"] = {"tabIds": [resolved_tab_id]}

    result = docs.documents().batchUpdate(
        documentId=document_id,
        body={"requests": [request]},
    ).execute()

    replies = result.get("replies", [{}])
    count = replies[0].get("replaceAllText", {}).get("occurrencesChanged", 0) if replies else 0
    return json.dumps({"status": "ok", "occurrences_replaced": count, "tabId": resolved_tab_id})


@mcp.tool()
def batch_update(ctx: Context, document_id: str, requests: list[dict], tab_id: str = "", tab_title: str = "") -> str:
    """Execute a batch update on a Google Doc using the full batchUpdate endpoint.
    This provides access to all batchUpdate operations including formatting, styling, and structural changes.

    When tab_id or tab_title is provided, any request containing a 'location' or 'range'
    without an explicit tabId will have the resolved tabId injected automatically.

    Args:
        document_id: The document ID
        requests: A list of request objects. Common operations include:
            - updateTextStyle: Apply bold, italic, underline, font, color to text ranges
            - updateParagraphStyle: Change alignment, spacing, indentation
            - insertText: Insert text at a position
            - deleteContentRange: Delete a range of content
            - insertInlineImage: Insert an image
            - createNamedRange: Create a named range
            - insertTable: Insert a table
            - insertTableRow / insertTableColumn: Modify tables
            - updateTableCellStyle: Style table cells

            Example requests:
            [
                {
                    "updateTextStyle": {
                        "range": {"startIndex": 1, "endIndex": 10},
                        "textStyle": {"bold": true, "foregroundColor": {"color": {"rgbColor": {"red": 0.2, "green": 0.4, "blue": 0.9}}}},
                        "fields": "bold,foregroundColor"
                    }
                },
                {
                    "updateParagraphStyle": {
                        "range": {"startIndex": 1, "endIndex": 10},
                        "paragraphStyle": {"namedStyleType": "HEADING_1"},
                        "fields": "namedStyleType"
                    }
                }
            ]
        tab_id: Optional tab ID — auto-injected into ranges/locations missing a tabId
        tab_title: Optional tab title (case-insensitive) — resolved to tab ID
    """
    docs, _ = _get_ctx(ctx)

    # Resolve tab
    resolved_tab_id = tab_id
    if not resolved_tab_id and tab_title:
        doc = _get_doc_with_tabs(docs, document_id)
        _, resolved_tab_id = _resolve_tab(doc, tab_title=tab_title)

    # Inject tabId into requests that have location/range without one
    if resolved_tab_id:
        for req in requests:
            for op_name, op_body in req.items():
                if not isinstance(op_body, dict):
                    continue
                for key in ("location", "range", "insertionLocation"):
                    if key in op_body and "tabId" not in op_body[key]:
                        op_body[key]["tabId"] = resolved_tab_id
                # Also handle nested range inside replaceAllText etc.
                if "containsText" in op_body and "tabsCriteria" not in op_body:
                    op_body["tabsCriteria"] = {"tabIds": [resolved_tab_id]}

    result = docs.documents().batchUpdate(
        documentId=document_id,
        body={"requests": requests},
    ).execute()

    replies = result.get("replies", [])
    return json.dumps({
        "status": "ok",
        "replies_count": len(replies),
        "documentId": result.get("documentId", document_id),
        "tabId": resolved_tab_id,
    }, indent=2)


@mcp.tool()
def list_documents(ctx: Context, query: str = "", max_results: int = 20) -> str:
    """Search for Google Docs in Drive by name.

    Args:
        query: Search term for document names (empty = list recent docs)
        max_results: Max number of results (default 20, max 100)
    """
    _, drive = _get_ctx(ctx)
    max_results = min(max_results, 100)

    q = "mimeType='application/vnd.google-apps.document' and trashed=false"
    if query:
        safe_query = query.replace("'", "\\'")
        q += f" and name contains '{safe_query}'"

    results = (
        drive.files()
        .list(
            q=q,
            pageSize=max_results,
            fields="files(id, name, modifiedTime, webViewLink)",
            orderBy="modifiedTime desc",
        )
        .execute()
    )

    files = results.get("files", [])
    if not files:
        return json.dumps({"documents": [], "message": "No documents found"})

    docs_list = [
        {
            "documentId": f["id"],
            "title": f["name"],
            "modifiedTime": f.get("modifiedTime", ""),
            "link": f.get("webViewLink", ""),
        }
        for f in files
    ]
    return json.dumps({"documents": docs_list, "count": len(docs_list)}, indent=2)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
