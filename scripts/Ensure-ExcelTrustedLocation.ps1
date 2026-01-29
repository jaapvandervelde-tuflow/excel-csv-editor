param(
    [Parameter(Mandatory=$true)]
    [string]$TrustedPath,

    [switch]$AllowSubfolders,

    [string]$Description = "CSV Editor temp (auto-added)",

    [switch]$Prompt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertTo-CanonicalPath([string]$p) {
    $rp = (Resolve-Path -LiteralPath $p).Path
    if (-not $rp.EndsWith("\")) { $rp += "\" }
    return $rp
}

function Test-TrustedLocationExists([string]$tlRoot, [string]$pathNorm) {
    if (-not (Test-Path -LiteralPath $tlRoot)) { return $false }

    Get-ChildItem -LiteralPath $tlRoot -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $p = (Get-ItemPropertyValue -LiteralPath $_.PSPath -Name Path -ErrorAction Stop)
            if ($p) {
                if (-not $p.EndsWith("\")) { $p += "\" }
                if ($p -ieq $pathNorm) { return $true }
            }
        } catch { }
    }
    return $false
}

function Add-TrustedLocation([string]$tlRoot, [string]$pathNorm) {
    if (-not (Test-Path -LiteralPath $tlRoot)) {
        New-Item -Path $tlRoot -Force | Out-Null
    }

    $i = 1
    while (Test-Path -LiteralPath (Join-Path $tlRoot ("Location{0}" -f $i))) { $i++ }
    $loc = Join-Path $tlRoot ("Location{0}" -f $i)

    New-Item -Path $loc -Force | Out-Null
    New-ItemProperty -Path $loc -Name Path -Value $pathNorm -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $loc -Name AllowSubfolders -Value ([int]$AllowSubfolders.IsPresent) -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $loc -Name Description -Value $Description -PropertyType String -Force | Out-Null
}

$trustedNorm = ConvertTo-CanonicalPath $TrustedPath

$officeRoot = "HKCU:\Software\Microsoft\Office"
if (-not (Test-Path -LiteralPath $officeRoot)) {
    # Excel not installed for this user profile, or no Office HKCU hive yet
    exit 0
}

# Candidate version keys under HKCU\Software\Microsoft\Office\<ver>\Excel\Security\Trusted Locations
$versions = Get-ChildItem -LiteralPath $officeRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+(\.\d+)?$' } |
    ForEach-Object { $_.PSChildName } |
    Sort-Object -Descending

$targets = @()
foreach ($ver in $versions) {
    $tlRoot = Join-Path $officeRoot "$ver\Excel\Security\Trusted Locations"
    # Only act on versions where Excel security key exists or likely exists
    $targets += $tlRoot
}

if ($targets.Count -eq 0) { exit 0 }

# Idempotency: if any version already has it, do nothing.
foreach ($tl in $targets) {
    if (Test-TrustedLocationExists $tl $trustedNorm) { exit 0 }
}

if ($Prompt) {
    Write-Host "Excel Trusted Location is required for macros to run from:"
    Write-Host "  $trustedNorm"
    $resp = Read-Host "Add this trusted location for the current user? (Y/N)"
    if ($resp -notin @("Y","y")) { exit 2 }
}

foreach ($tl in $targets) {
    try {
        Add-TrustedLocation $tl $trustedNorm
    } catch {
        # Some versions may not allow writes; ignore and continue
    }
}

exit 0
