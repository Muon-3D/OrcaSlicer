<# 
MUON-PROFILE-UPDATE-SCRIPT-GENERATOR.ps1

One-click generator:
- Clears output dir
- Uses timestamp version automatically
- Outputs ONLY the updater .ps1 (no .cmd, no .zip)

Payload contains:
  - Muon3D\ folder
  - Muon3D.json
#>

[CmdletBinding()]
param(
    [string]$SourceProfilesDir = "D:\Dropbox\Muon3D_SharedFolder\Slicing\OrcaSlicer\OrcaSlicer\resources\profiles",
    [string]$OutputDir = (Join-Path $PSScriptRoot "build\MUON-DEV-profile-updater"),
    [string]$VendorName = "Muon3D"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Timestamp version (local time)
$Version = Get-Date -Format "yyyy.MM.dd.HHmmss"

function Reset-Dir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
        return
    }

    # Try to delete contents, not the folder itself (avoids common locks)
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | ForEach-Object {
                try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop }
                catch { throw }
            }
            return
        }
        catch {
            if ($attempt -eq 6) { throw }
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }
}

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function New-ZipFromFolder {
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$ZipPath
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $ZipPath) {
        Remove-Item -LiteralPath $ZipPath -Force
    }

    # IMPORTANT: ZipPath must not be inside SourceFolder
    [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceFolder, $ZipPath)
}

# Clear output dir each run
Reset-Dir $OutputDir

$vendorFolder = Join-Path $SourceProfilesDir $VendorName
$vendorJson   = Join-Path $SourceProfilesDir ("{0}.json" -f $VendorName)

if (-not (Test-Path -LiteralPath $vendorFolder)) { throw "Missing vendor folder: $vendorFolder" }
if (-not (Test-Path -LiteralPath $vendorJson))   { throw "Missing vendor json:   $vendorJson" }

# Stage payload in temp
$stagingRoot = Join-Path $env:TEMP ("{0}_UpdaterBuild_{1}" -f $VendorName, ([Guid]::NewGuid().ToString("N")))
Ensure-Dir $stagingRoot

# Zip temp file (outside stagingRoot)
$zipTempPath = Join-Path $env:TEMP ("{0}_{1}.zip" -f $VendorName, ([Guid]::NewGuid().ToString("N")))

try {
    Copy-Item -LiteralPath $vendorFolder -Destination (Join-Path $stagingRoot $VendorName) -Recurse -Force
    Copy-Item -LiteralPath $vendorJson   -Destination (Join-Path $stagingRoot ("{0}.json" -f $VendorName)) -Force

    Write-Verbose "StagingRoot: $stagingRoot"
    Write-Verbose "OutputDir:   $OutputDir"
    Write-Verbose "Version:     $Version"

    $stagedItems = Get-ChildItem -LiteralPath $stagingRoot -Force
    Write-Verbose ("Items staged: " + ($stagedItems.Name -join ", "))

    # Create temp ZIP and base64 it (zip is NOT written to output dir)
    New-ZipFromFolder -SourceFolder $stagingRoot -ZipPath $zipTempPath
    if (-not (Test-Path -LiteralPath $zipTempPath)) {
        throw "ZIP was not created at: $zipTempPath"
    }

    $zipBytes = [System.IO.File]::ReadAllBytes($zipTempPath)
    $b64 = [Convert]::ToBase64String($zipBytes)

    $sb = New-Object System.Text.StringBuilder
    for ($i=0; $i -lt $b64.Length; $i += 120) {
        $len = [Math]::Min(120, $b64.Length - $i)
        [void]$sb.AppendLine($b64.Substring($i, $len))
    }
    $b64Wrapped = $sb.ToString().TrimEnd()

    # Updater script template
    $template = @"
<# 
Update-$VendorName-OrcaSlicerProfiles.ps1
Single-file updater. Replaces $VendorName folder + $VendorName.json
in:
  - Program Files: <OrcaInstallDir>\resources\profiles
  - Roaming: %AppData%\OrcaSlicer\system\profiles (or \system if \profiles doesn't exist)

Right-click -> Run with PowerShell.
#>

[CmdletBinding()]
param(
    [switch]`$ForceCloseOrcaSlicer = `$true,
    [string]`$OrcaInstallDir,
    [string]`$RoamingSystemDir,
    [switch]`$SkipProgramFiles,
    [switch]`$SkipRoaming
)

Set-StrictMode -Version Latest
`$ErrorActionPreference = "Stop"

trap {
    Write-Host ""
    Write-Host "ERROR: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close"
    break
}

`$VendorName = "$VendorName"

`$PayloadZipB64 = @'
$b64Wrapped
'@

function Test-IsAdmin {
    `$id = [Security.Principal.WindowsIdentity]::GetCurrent()
    `$p  = New-Object Security.Principal.WindowsPrincipal(`$id)
    return `$p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Admin {
    if (Test-IsAdmin) { return }
    Write-Host "UAC required (Program Files). Relaunching as Administrator..." -ForegroundColor Yellow

    # Pass args as array (no fragile quote escaping)
    `$argList = @("-NoProfile","-ExecutionPolicy","Bypass","-File", `$PSCommandPath)
    if (`$MyInvocation.UnboundArguments) { `$argList += `$MyInvocation.UnboundArguments }

    Start-Process -FilePath "powershell.exe" -ArgumentList `$argList -Verb RunAs | Out-Null
    exit 0
}

function Ensure-Dir([string]`$Path) {
    if (-not (Test-Path -LiteralPath `$Path)) {
        New-Item -ItemType Directory -Path `$Path | Out-Null
    }
}

function Get-OrcaInstallDirAuto {
    `$candidates = @()
    if (`$env:ProgramFiles) { `$candidates += (Join-Path `$env:ProgramFiles "OrcaSlicer") }
    if (`${env:ProgramFiles(x86)}) { `$candidates += (Join-Path `${env:ProgramFiles(x86)} "OrcaSlicer") }

    foreach (`$c in `$candidates) {
        if (Test-Path -LiteralPath `$c) { return `$c }
    }
    return `$null
}

function Stop-OrcaSlicerIfRunning {
    `$procs = Get-Process -Name "OrcaSlicer" -ErrorAction SilentlyContinue
    if (-not `$procs) { return }

    Write-Host "OrcaSlicer is running. Attempting to close..." -ForegroundColor Yellow

    foreach (`$p in `$procs) {
        try {
            if (`$p.MainWindowHandle -ne 0) { [void]`$p.CloseMainWindow() }
        } catch {}
    }

    Start-Sleep -Seconds 2
    `$still = Get-Process -Name "OrcaSlicer" -ErrorAction SilentlyContinue
    if (-not `$still) { return }

    if (-not `$ForceCloseOrcaSlicer) {
        throw "OrcaSlicer is still running. Close it and re-run, or use -ForceCloseOrcaSlicer."
    }

    Write-Host "Force-killing OrcaSlicer..." -ForegroundColor Yellow
    foreach (`$p in `$still) {
        try { Stop-Process -Id `$p.Id -Force -ErrorAction Stop } catch {}
    }
}

function Remove-IfExists([string]`$Path) {
    if (Test-Path -LiteralPath `$Path) {
        Remove-Item -LiteralPath `$Path -Recurse -Force
    }
}

function Install-ToBase([string]`$TargetBase, [string]`$ExtractRoot, [string]`$Label) {
    if (`$TargetBase -notmatch "OrcaSlicer") {
        throw "Refusing to operate on base path not containing 'OrcaSlicer': `$TargetBase"
    }

    Ensure-Dir `$TargetBase

    `$targetFolder = Join-Path `$TargetBase `$VendorName
    `$targetJson   = Join-Path `$TargetBase ("{0}.json" -f `$VendorName)

    `$srcFolder = Join-Path `$ExtractRoot `$VendorName
    `$srcJson   = Join-Path `$ExtractRoot ("{0}.json" -f `$VendorName)

    if (-not (Test-Path -LiteralPath `$srcFolder)) { throw "Payload missing folder: `$srcFolder" }
    if (-not (Test-Path -LiteralPath `$srcJson))   { throw "Payload missing file:   `$srcJson" }

    Write-Host ""
    Write-Host ("Installing to: " + `$Label) -ForegroundColor Cyan
    Write-Host ("Base: " + `$TargetBase)

    Remove-IfExists `$targetFolder
    Remove-IfExists `$targetJson

    Copy-Item -LiteralPath `$srcFolder -Destination `$targetFolder -Recurse -Force
    Copy-Item -LiteralPath `$srcJson   -Destination `$targetJson -Force
}

# --- Main ---
# Detect install dirs
if (-not `$OrcaInstallDir) { `$OrcaInstallDir = Get-OrcaInstallDirAuto }
if (-not `$RoamingSystemDir) { `$RoamingSystemDir = Join-Path `$env:APPDATA "OrcaSlicer\system" }

# Only elevate if we are actually going to touch Program Files AND we detected an install dir
if (-not `$SkipProgramFiles -and -not [string]::IsNullOrWhiteSpace(`$OrcaInstallDir)) {
    Ensure-Admin
}

Stop-OrcaSlicerIfRunning

`$stagingRoot  = Join-Path `$env:TEMP ("{0}_OrcaUpdater_{1}" -f `$VendorName, ([Guid]::NewGuid().ToString("N")))
`$zipPath      = Join-Path `$stagingRoot "payload.zip"
`$extractRoot  = Join-Path `$stagingRoot "extracted"

Ensure-Dir `$stagingRoot
Ensure-Dir `$extractRoot

try {
    `$b64 = (`$PayloadZipB64 -replace "\s","")
    if ([string]::IsNullOrWhiteSpace(`$b64)) { throw "Embedded payload is empty." }

    `$bytes = [Convert]::FromBase64String(`$b64)
    [System.IO.File]::WriteAllBytes(`$zipPath, `$bytes)

    Expand-Archive -LiteralPath `$zipPath -DestinationPath `$extractRoot -Force

    # Program Files update is best-effort
    if (-not `$SkipProgramFiles) {
        if ([string]::IsNullOrWhiteSpace(`$OrcaInstallDir)) {
            Write-Host "WARNING: Could not auto-detect OrcaSlicer install dir (Program Files). Skipping Program Files update." -ForegroundColor Yellow
        } else {
            `$pfBase = Join-Path `$OrcaInstallDir "resources\profiles"
            if (-not (Test-Path -LiteralPath `$pfBase)) {
                Write-Host "WARNING: OrcaSlicer profiles dir not found at: `$pfBase (skipping Program Files update)" -ForegroundColor Yellow
            } else {
                Install-ToBase -TargetBase `$pfBase -ExtractRoot `$extractRoot -Label "Program Files (resources\profiles)"
            }
        }
    }

    if (-not `$SkipRoaming) {
        `$candidateA = Join-Path `$RoamingSystemDir "profiles"
        `$roamBase = if (Test-Path -LiteralPath `$candidateA) { `$candidateA } else { `$RoamingSystemDir }
        Ensure-Dir `$roamBase
        Install-ToBase -TargetBase `$roamBase -ExtractRoot `$extractRoot -Label "Roaming (%AppData%\OrcaSlicer\system...)"
    }

    Write-Host ""
    Write-Host "SUCCESS: Updated profiles." -ForegroundColor Green
}
finally {
    try { Remove-Item -LiteralPath `$stagingRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}

Write-Host ""
Read-Host "Press Enter to close"
"@

    $updaterName = "Update-{0}-OrcaSlicerProfiles-{1}.ps1" -f $VendorName, $Version
    $updaterPath = Join-Path $OutputDir $updaterName
    Set-Content -LiteralPath $updaterPath -Value $template -Encoding UTF8

    Write-Host ""
    Write-Host "Build complete:" -ForegroundColor Green
    Write-Host "Updater: $updaterPath"
}
finally {
    # Clean temp artifacts
    try { if (Test-Path -LiteralPath $zipTempPath) { Remove-Item -LiteralPath $zipTempPath -Force } } catch {}
    try { Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}
