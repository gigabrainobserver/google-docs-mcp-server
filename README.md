# Google Docs MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io/) server for Google Docs with **full tab support** — something most existing servers get wrong or skip entirely.

Uses Google service accounts with domain-wide delegation, so it works in Workspace orgs without OAuth consent screens.

<!-- mcp-name: io.github.gigabrainobserver/google-docs-mcp-server -->

## What it does

| Tool | Description |
|------|-------------|
| `list_tabs` | List all tabs (including nested) in a document |
| `read_document` | Read one tab or all tabs as markdown with proper heading levels |
| `get_document_info` | Lightweight metadata: title, ID, link, tab list |
| `create_document` | Create a new doc with optional initial text |
| `append_text` | Append text to the end of a specific tab |
| `insert_text` | Insert text at a character index in a tab |
| `replace_text` | Find and replace within a tab (or whole doc) |
| `batch_update` | Full batchUpdate access — formatting, tables, images, styles |
| `list_documents` | Search Drive for docs by name |

All tab-targeting tools accept `tab_id` or `tab_title` (case-insensitive). The `batch_update` tool auto-injects `tabId` into requests so you don't have to.

## Why this exists

Google Docs has supported [tabs](https://workspaceupdates.googleblog.com/2024/10/google-docs-tabs.html) since late 2024, but most MCP servers either:
- Ignore tabs entirely (only read the first tab)
- Don't use `includeTabsContent=True`, so tab content is invisible
- Don't handle nested tabs

This server handles all of that correctly and converts content to markdown with proper heading structure.

## Install

```bash
pip install google-docs-mcp-server
```

Or run directly with [uv](https://docs.astral.sh/uv/):

```bash
uvx google-docs-mcp-server
```

## Prerequisites

- Python 3.11+
- A Google Cloud service account with domain-wide delegation

## Setup

### 1. Create a GCP service account

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create (or select) a project
2. Enable the **Google Docs API** and **Google Drive API**
3. Create a service account under **IAM & Admin > Service Accounts**
4. Create a JSON key and download it

### 2. Enable domain-wide delegation

1. In GCP, on the service account details page, enable **Domain-wide Delegation** and note the Client ID
2. In [Google Workspace Admin](https://admin.google.com/) > Security > API Controls > Domain-wide Delegation
3. Add the Client ID with these scopes:
   ```
   https://www.googleapis.com/auth/documents
   https://www.googleapis.com/auth/drive.readonly
   ```

### 3. Configure your MCP client

Add to your MCP config (e.g. `~/.claude/mcp.json` or `.mcp.json`):

```json
{
  "mcpServers": {
    "google-docs": {
      "command": "uvx",
      "args": ["google-docs-mcp-server"],
      "env": {
        "SERVICE_ACCOUNT_PATH": "/path/to/your-service-account-key.json",
        "SUBJECT_EMAIL": "you@yourdomain.com"
      }
    }
  }
}
```

`SUBJECT_EMAIL` is the Workspace user the service account impersonates.

## Environment variables

| Variable | Required | Description |
|----------|----------|-------------|
| `SERVICE_ACCOUNT_PATH` | Yes | Path to the service account JSON key file |
| `SUBJECT_EMAIL` | Yes | Email of the Workspace user to impersonate |

## License

MIT
