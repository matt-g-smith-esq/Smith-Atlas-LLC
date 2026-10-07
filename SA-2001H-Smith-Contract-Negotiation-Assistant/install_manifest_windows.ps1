# Copyright 2026 Smith-Atlas LLC (SA-2001H-V1). MIT License.
#
# Windows sideload installer (for LOCAL TESTING and SINGLE-MACHINE installs).
#
# Windows does not use a simple file-copy like macOS. Word loads sideloaded add-ins
# from a "Trusted Add-in Catalog" folder that you register once in Word's Trust Center.
# This script does the scriptable parts (stamp the manifest, place it in a catalog
# folder) and then prints the one-time manual Trust Center steps you complete in Word.
#
# We intentionally use the MANUAL Trust Center method (no registry edits) so nothing
# touches the Windows registry.
#
# For organization-wide rollout, use centralized deployment via the Microsoft 365
# admin center instead (see the user guide) — that path is OS-agnostic and does not
# use this script.
#
# Run from PowerShell:   powershell -ExecutionPolicy Bypass -File .\install_manifest_windows.ps1

$ErrorActionPreference = "Stop"

Write-Host "====================================================="
Write-Host "  Contract Negotiation Assistant - Windows Sideload  "
Write-Host "====================================================="

# Resolve folders relative to this script (expected to live in Contract-Frontend\).
$FrontendDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackendDir  = Join-Path (Split-Path -Parent $FrontendDir) "Contract-Backend"

# --- Step 1: Stamp the server URL from .env into manifest.xml (and taskpane.js) ---
$SetupScript = Join-Path $BackendDir "setup.js"
if (Test-Path $SetupScript) {
    Write-Host "Stamping server URL from .env into the manifest..."
    Push-Location $BackendDir
    & node setup.js
    $stampExit = $LASTEXITCODE
    Pop-Location
    if ($stampExit -ne 0) {
        Write-Host ""
        Write-Host "CRITICAL ERROR: Manifest stamping failed. Fix the error above" -ForegroundColor Red
        Write-Host "(usually NODE_SERVER_URL in .env), then re-run this script." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit 1
    }
} else {
    Write-Host "WARNING: setup.js not found in Contract-Backend; using existing manifest.xml as-is." -ForegroundColor Yellow
}

$ManifestPath = Join-Path $FrontendDir "manifest.xml"
if (-not (Test-Path $ManifestPath)) {
    Write-Host "CRITICAL ERROR: manifest.xml was not found after stamping." -ForegroundColor Red
    Write-Host "Expected path: $ManifestPath" -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}

# --- Step 2: Place the manifest in a local catalog folder ---
# Word requires the catalog to be a *shared* folder (a UNC path). We create a folder
# and share it for the current user. The user then registers its network path in
# Word's Trust Center (manual step printed below).
$CatalogDir = Join-Path $env:USERPROFILE "SmithAtlasAddInCatalog"
if (-not (Test-Path $CatalogDir)) {
    New-Item -ItemType Directory -Path $CatalogDir | Out-Null
    Write-Host "Created add-in catalog folder: $CatalogDir"
}
Copy-Item -Path $ManifestPath -Destination (Join-Path $CatalogDir "manifest.xml") -Force
Write-Host "Placed manifest into the catalog folder."

# Try to share the folder so it has a UNC network path (required by Trust Center).
# This may require an elevated prompt; if it fails, the user can share it manually.
$ShareName = "SmithAtlasAddInCatalog"
$networkPath = "\\$env:COMPUTERNAME\$ShareName"
try {
    $existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-SmbShare -Name $ShareName -Path $CatalogDir -ReadAccess "$env:USERNAME" -ErrorAction Stop | Out-Null
        Write-Host "Shared the catalog folder as: $networkPath"
    } else {
        Write-Host "Catalog folder already shared as: $networkPath"
    }
} catch {
    Write-Host ""
    Write-Host "NOTE: Could not automatically share the folder (this often needs administrator rights)." -ForegroundColor Yellow
    Write-Host "You can share it manually: right-click '$CatalogDir' > Properties > Sharing > Share." -ForegroundColor Yellow
    Write-Host "After sharing, note the network path (looks like \\$env:COMPUTERNAME\\$ShareName)." -ForegroundColor Yellow
    $networkPath = "(the network path shown after you share the folder)"
}

# --- Step 3: Print the one-time manual Trust Center steps ---
Write-Host ""
Write-Host "-----------------------------------------------------------------"
Write-Host " ONE-TIME SETUP IN WORD (Trust Center)" -ForegroundColor Cyan
Write-Host "-----------------------------------------------------------------"
Write-Host " 1. Open Word. Go to:  File > Options > Trust Center >"
Write-Host "    Trust Center Settings > Trusted Add-in Catalogs"
Write-Host " 2. In 'Catalog Url', paste this network path:"
Write-Host "       $networkPath"
Write-Host " 3. Click 'Add catalog'. Check 'Show in Menu'. Click OK, then OK."
Write-Host " 4. Close and reopen Word."
Write-Host " 5. On the Home tab choose Add-ins (or Insert > Get Add-ins) >"
Write-Host "    SHARED FOLDER, then choose 'Contract Negotiation Assistant'."
Write-Host "-----------------------------------------------------------------"
Write-Host ""
Write-Host "You only do the Trust Center step once. After that, re-running this"
Write-Host "script (e.g. after changing your server URL) just refreshes the manifest."
Write-Host ""
Read-Host "Press Enter to close"
