# acn-m365-local — CLAUDE.md

Questo è il server MCP locale che simula il connettore Microsoft 365 di Claude.ai,
leggendo direttamente dai file OneDrive sincronizzati. Vedi README.md per il contesto.

---

## Permessi automatici (nessun prompt richiesto)

### Lettura file
- Tutti i file in questa repo
- Tutti i file sotto `C:\Users\paola.loi\OneDrive - Accenture\`

### Bash
- `python server.py` — avvio/test del server MCP
- `python -m pip install ...` / `pip install -r requirements.txt`
- `python --version`, `pip --version`
- `mcp run server.py`
- `git status`, `git log`, `git diff`, `git show`, `git branch`
- `claude mcp list`, `claude mcp add`, `claude mcp remove`

### PowerShell
- `python` e `pip` (qualsiasi variante)
- `Get-ChildItem`, `Get-Content`, `Test-Path`, `Resolve-Path`
- `Select-Object`, `Where-Object`, `Sort-Object`, `Format-Table`
- `.\install.ps1`
- `claude mcp list`, `claude mcp add`, `claude mcp remove`
- `Set-Content`, `Out-File` (solo nella repo o in OneDrive - Accenture)

### MCP tools (acn-m365)
- `list_project_files`
- `read_document`
- `search_documents`
- `get_recent_files`
- `create_document`
- `update_document`
- `create_summary_report`

---

## Operazioni che richiedono conferma

- `git commit`, `git push`, `git reset`, `git checkout`
- Eliminazione di file (`Remove-Item`, `rm`, `del`)
- Scrittura fuori da questa repo o da `OneDrive - Accenture`
- Modifica di `.claude/settings.json` o `.claude/settings.local.json`

---

## Note di sviluppo

- Entry point: [server.py](server.py) — variabile `mcp` (istanza FastMCP)
- Trasporto: stdio (default Claude Code MCP)
- OneDrive base: `C:\Users\{USERNAME}\OneDrive - Accenture\` (USERNAME da `os.environ`)
- Tutti i path validati contro la base dir (protezione path traversal)
- Librerie documento importate lazy (server parte anche senza dipendenze installate)
