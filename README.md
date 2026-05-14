# acn-m365-local

A Python MCP server that simulates Claude.ai's Microsoft 365 connector by
reading directly from locally synced OneDrive files.

**Zero OAuth. Zero IT approval. 100% local.**

---

## Prerequisites

- Python 3.10+
- Claude Code CLI (`npm install -g @anthropic-ai/claude-code`)
- OneDrive for Business synced to `C:\Users\{USERNAME}\OneDrive - Accenture\`

---

## Installation (3 steps)

```powershell
git clone https://github.com/paola-loi/acn-m365-local.git
cd acn-m365-local
.\install.ps1
```

The script:
1. Verifies Python is installed
2. Installs the dependencies (`mcp`, `python-docx`, `openpyxl`, `python-pptx`)
3. Registers the MCP server with Claude Code
4. Confirms with `claude mcp list`

After installation, **restart Claude Code** and verify with `/mcp`.

---

## Expected OneDrive layout

```
C:\Users\{USERNAME}\OneDrive - Accenture\
├── MARS\
│   ├── kickoff.docx
│   ├── budget_2025.xlsx
│   └── roadmap.pptx
├── Unilever\
│   ├── meeting_notes.md
│   └── deliverables.docx
└── ...
```

The server treats any subfolder as a "project".

---

## Supported formats

| Extension | Library |
|-----------|---------|
| `.docx` | python-docx |
| `.xlsx` | openpyxl |
| `.pptx` | python-pptx |
| `.md` / `.txt` | stdlib |

---

## Usage examples

```
"List all files in the MARS folder"
"Show me only the Excel files in the Unilever project"
"Read the file MARS/budget_2025.xlsx"
"Search 'action items' across all documents"
"Search 'forecast' only in the MARS project"
"Which files did I modify this week?"
"Create a notes file at MARS/meeting_20260505.md"
"Generate a summary report of the MARS project"
"List the sheets of MARS/budget_2025.xlsx"
"Read range A1:D50 of the Forecast sheet in MARS/budget_2025.xlsx"
"Explain the formula in cell C12 of MARS/budget_2025.xlsx"
"Add a Margin = Revenue - Cost column in budget_2025.xlsx"
"Find where 'forecast' appears in the formulas of budget_2025.xlsx"
```

> Writes to xlsx files **never overwrite** the original: they are saved
> alongside as `<name>.claude.xlsx`. Open that file in Excel to see the
> calculated formulas (openpyxl only writes the expression, not the value).

---

## Available tools

| Tool | Description |
|------|-------------|
| `list_project_files` | List files in a folder, with optional extension filter |
| `read_document` | Read the contents of a file |
| `search_documents` | Search a keyword across the contents of all files |
| `get_recent_files` | Files modified in the last N days |
| `create_document` | Create a new `.md` or `.docx` file |
| `update_document` | Overwrite an existing file |
| `create_summary_report` | Generate an aggregated `.md` report of a project |
| `list_sheets` | List sheets of an xlsx with dimensions |
| `describe_sheet` | Headers, inferred column types, named ranges of a sheet |
| `read_cells` | Read an A1 range with values **and** formulas |
| `write_cells` | Apply cell changes, saved to a `<name>.claude.xlsx` copy |
| `find_in_xlsx` | Find a string across values and formulas of an xlsx |
| `view_xlsx` | Return the URL of a live HTML viewer (auto-reload on write) |

---

## Live viewer (experimental)

When the server starts, a small HTTP server runs on `http://127.0.0.1:8765`
(stdlib, zero extra dependencies). Bound to localhost only.

Ask Claude *"open the viewer for MARS/budget.claude.xlsx"* → you get the URL.
The page shows a table for each sheet and reloads automatically:

- when Claude writes via `write_cells`
- when you save the file from LibreOffice/Excel (mtime polling ~1s)

Limits:
- it is a **viewer**, not an editor — you can't click on cells.
- formulas display the last *cached* value from Excel (`openpyxl` does not
  recalculate). Open the file in Excel/Calc to see real calculations.
- truncated to 1000 rows × 100 columns per sheet (banner shown if truncated).
- if port 8765 is busy the server tries 8766–8775, then a random port —
  check the process stderr for the actual port.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| `OneDrive base path not found` | Make sure OneDrive is running and synced |
| `File not found` | Check the path and wait for OneDrive sync to finish |
| `Permission denied` | Close the file in Office and wait for sync |
| `Module not found` | Run `pip install -r requirements.txt` again |
| Server doesn't appear in `/mcp` | Run `claude mcp list` and restart Claude Code |

---

## Local development

```powershell
# Test server startup (waits for JSON-RPC on stdin — Ctrl+C to exit)
python server.py

# Or via the MCP CLI
mcp run server.py
```
