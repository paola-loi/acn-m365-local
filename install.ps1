#Requires -Version 5.1
<#
.SYNOPSIS
    Installs the acn-m365-local MCP server into Claude Code.
.DESCRIPTION
    1. Verifies Python is available.
    2. Installs Python dependencies from requirements.txt.
    3. Registers the MCP server with Claude Code (project scope).
    4. Confirms installation with claude mcp list.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot   = $PSScriptRoot
$ServerPath = Join-Path $RepoRoot "server.py"
$ReqPath    = Join-Path $RepoRoot "requirements.txt"

Write-Host ""
Write-Host "=== acn-m365-local installer ===" -ForegroundColor Cyan

# 1 — Verify Python ───────────────────────────────────────────────
Write-Host ""
Write-Host "[1/4] Checking Python..." -ForegroundColor Yellow
try {
    $pyVersion = python --version 2>&1
    Write-Host "  Found: $pyVersion" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "ERROR: Python not found in PATH." -ForegroundColor Red
    Write-Host "Install Python 3.10+ from https://python.org and re-run this script." -ForegroundColor Red
    exit 1
}

# 2 — Install dependencies ────────────────────────────────────────
Write-Host ""
Write-Host "[2/4] Installing Python dependencies..." -ForegroundColor Yellow
python -m pip install -r $ReqPath
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: pip install failed (exit code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "  Dependencies installed." -ForegroundColor Green

# 3 — Register MCP server with Claude Code ────────────────────────
Write-Host ""
Write-Host "[3/4] Registering MCP server with Claude Code..." -ForegroundColor Yellow

$registered = $false
try {
    claude mcp add acn-m365 python $ServerPath --scope project
    if ($LASTEXITCODE -eq 0) {
        $registered = $true
        Write-Host "  Registered via 'claude mcp add'." -ForegroundColor Green
    }
} catch {
    # fall through to manual fallback
}

if (-not $registered) {
    Write-Host "  'claude mcp add' unavailable — writing .claude.json manually..." -ForegroundColor Yellow
    $config = @{
        mcpServers = @{
            "acn-m365" = @{
                command = "python"
                args    = @($ServerPath)
            }
        }
    }
    $configPath = Join-Path $RepoRoot ".claude.json"
    $config | ConvertTo-Json -Depth 5 | Set-Content -Path $configPath -Encoding UTF8
    Write-Host "  Wrote $configPath" -ForegroundColor Green
}

# 4 — Confirm ─────────────────────────────────────────────────────
Write-Host ""
Write-Host "[4/4] Installed MCP servers:" -ForegroundColor Yellow
claude mcp list

Write-Host ""
Write-Host "=== Installation complete! ===" -ForegroundColor Cyan
Write-Host "Restart Claude Code, then type /mcp to verify the server is connected." -ForegroundColor White
Write-Host ""
