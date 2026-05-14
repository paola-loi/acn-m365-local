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
- `list_sheets`
- `describe_sheet`
- `read_cells`
- `write_cells`
- `find_in_xlsx`
- `view_xlsx`

---

## Workflow xlsx (editing iterativo)

Quando l'utente sta lavorando su un xlsx:

- Dopo **ogni** chiamata a `write_cells`, richiama subito `read_cells` sullo
  stesso range (o un range che include le celle toccate + 2 righe di contesto)
  e mostra il risultato in chat come tabella markdown. L'utente non deve
  aprire il file per vedere cos'è cambiato.
- Iterazioni successive sullo stesso file devono continuare a modificare la
  copia `<nome>.claude.xlsx` esistente (è il comportamento di default di
  `write_cells`: se la copia esiste, la aggiorna). Non rifare la copia da zero
  ogni turno.
- Prima della prima modifica su un xlsx mai visto in questa conversazione,
  chiama `list_sheets` e poi `read_cells` sul range rilevante per orientarti —
  non assumere la struttura.
- L'utente probabilmente tiene il file aperto in **LibreOffice Calc** accanto:
  ricordagli nel primo messaggio della sessione xlsx di fare `File → Reload`
  dopo ogni scrittura per vedere le formule calcolate. Excel di solito tiene
  il file in lock e impedirebbe la scrittura su `.claude.xlsx` se la copia è
  aperta lì.
- Se l'utente chiede di "vedere live" / "aprire nel browser" / un viewer,
  chiama `view_xlsx` e dagli l'URL — quella pagina si ricarica da sola sia
  sui tuoi `write_cells` sia sui salvataggi da LibreOffice. In quel caso non
  serve ricordare il Reload manuale; l'utente può anche dirti "non rileggere
  il range dopo ogni write, guardo dal browser" — rispetta la preferenza
  per il resto della sessione.

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
