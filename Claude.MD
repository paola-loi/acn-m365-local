# acn-m365-local — CLAUDE.md

This is the local MCP server that simulates Claude.ai's Microsoft 365
connector by reading directly from OneDrive-synced files. See README.md for
the broader context.

---

## Automatic permissions (no prompt required)

### File reads
- All files in this repo
- All files under `C:\Users\paola.loi\OneDrive - Accenture\`

### Bash
- `python server.py` — start/test the MCP server
- `python -m pip install ...` / `pip install -r requirements.txt`
- `python --version`, `pip --version`
- `mcp run server.py`
- `git status`, `git log`, `git diff`, `git show`, `git branch`
- `claude mcp list`, `claude mcp add`, `claude mcp remove`

### PowerShell
- `python` and `pip` (any variant)
- `Get-ChildItem`, `Get-Content`, `Test-Path`, `Resolve-Path`
- `Select-Object`, `Where-Object`, `Sort-Object`, `Format-Table`
- `.\install.ps1`
- `claude mcp list`, `claude mcp add`, `claude mcp remove`
- `Set-Content`, `Out-File` (only inside this repo or OneDrive - Accenture)

### MCP tools (acn-m365)
- `list_project_files`
- `read_document`
- `search_documents`
- `get_recent_files`
- `create_document`
- `update_document`
- `create_summary_report`
- `list_sheets`
- `describe_sheet`
- `read_cells`
- `write_cells`
- `find_in_xlsx`
- `view_xlsx`

---

## xlsx workflow (iterative editing)

When the user is working on an xlsx:

- After **every** `write_cells` call, immediately call `read_cells` on the
  same range (or a range that includes the touched cells + 2 rows of
  context) and show the result in chat as a markdown table. The user
  shouldn't have to open the file to see what changed.
- Subsequent iterations on the same file must keep modifying the existing
  `<name>.claude.xlsx` copy (this is `write_cells`'s default: if the copy
  exists, it is updated). Don't recreate the copy from scratch each turn.
- Before the first modification to an xlsx not yet seen in this conversation,
  call `list_sheets` and then `read_cells` on the relevant range to orient
  yourself — don't assume the structure.
- The user likely keeps the file open in **LibreOffice Calc** alongside:
  in the first message of an xlsx session, remind them to use
  `File → Reload` after each write to see the calculated formulas. Excel
  usually keeps the file locked and would block writes to `.claude.xlsx` if
  the copy is open there.
- If the user asks to "see it live" / "open in the browser" / wants a
  viewer, call `view_xlsx` and give them the URL — that page reloads on its
  own both for your `write_cells` calls and for LibreOffice saves. In that
  case you don't need to remind them about manual Reload; the user may also
  tell you "don't re-read the range after every write, I'm watching the
  browser" — honor that preference for the rest of the session.

---

## Operations that require confirmation

- `git commit`, `git push`, `git reset`, `git checkout`
- File deletion (`Remove-Item`, `rm`, `del`)
- Writes outside this repo or `OneDrive - Accenture`
- Changes to `.claude/settings.json` or `.claude/settings.local.json`

---

## Development notes

- Entry point: [server.py](server.py) — `mcp` variable (FastMCP instance)
- Transport: stdio (Claude Code MCP default)
- OneDrive base: `C:\Users\{USERNAME}\OneDrive - Accenture\` (USERNAME from `os.environ`)
- All paths validated against the base dir (path traversal protection)
- Document libraries imported lazily (the server can start without deps installed)
