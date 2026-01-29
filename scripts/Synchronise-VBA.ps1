param(
    # Defaults to ./CsvEditor.xlsm relative to the current directory
    [string]$WorkbookPath = ".\CsvEditor.xlsm",

    # Defaults to ../vba relative to this script
    [string]$OutDir = (Join-Path $PSScriptRoot "../vba"),

    # If set, do not delete existing exports before writing new ones
    [switch]$NoClean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-FullPath([string]$p) {
    $p = $p.Trim().Trim('"')
    return [System.IO.Path]::GetFullPath($p)
}

function Write-Text([string]$path, [string]$text) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $text, $utf8NoBom)
}

function Remove-IfExists([System.IO.DirectoryInfo]$dir, [string]$pattern) {
    Get-ChildItem -LiteralPath $dir.FullName -Filter $pattern -File -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Write-Step([string]$msg) {
    # Minimal, single-line status
    Write-Host ("[Synchronise-VBA] {0}" -f $msg)
}

$WorkbookPath = Get-FullPath $WorkbookPath
$OutDir = Get-FullPath $OutDir

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$dir = Get-Item -LiteralPath $OutDir

# VBComponent types (constants)
$vbext_ct_StdModule   = 1
$vbext_ct_ClassModule = 2
$vbext_ct_MSForm      = 3
$vbext_ct_Document    = 100

$excel = $null
$wb = $null

$exported = 0
$start = Get-Date

try {
    Write-Step "opening workbook: $WorkbookPath"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($WorkbookPath, $null, $false)

    # Optional: suppress UI in workbook code paths (if SetDevMode exists)
    try {
        $excel.Run("'" + $wb.Name + "'!ThisWorkbook.SetDevMode", $true)
        Write-Step "dev mode enabled"
    }
    catch {
        Write-Step "dev mode flag not set (macro not found or macros disabled)"
    }

    $vbproj = $wb.VBProject
    $components = $vbproj.VBComponents

    if (-not $NoClean) {
        Write-Step "cleaning output folder: $OutDir"
        Remove-IfExists $dir "*.bas"
        Remove-IfExists $dir "*.cls"
        Remove-IfExists $dir "*.frm"
        Remove-IfExists $dir "*.frx"
        Remove-IfExists $dir "_manifest.tsv"
    }
    else {
        Write-Step "skipping clean (NoClean)"
    }

    Write-Step ("exporting {0} VBA components..." -f $components.Count)

    for ($i = 1; $i -le $components.Count; $i++) {
        $c = $components.Item($i)

        $name = $c.Name
        $type = $c.Type

        $ext =
            if ($type -eq $vbext_ct_StdModule) { ".bas" }
            elseif ($type -eq $vbext_ct_ClassModule) { ".cls" }
            elseif ($type -eq $vbext_ct_MSForm) { ".frm" }
            elseif ($type -eq $vbext_ct_Document) { ".cls" }
            else { ".txt" }

        $outPath = Join-Path $dir.FullName ($name + $ext)
        $c.Export($outPath)

        $exported++

        # Minimal progress: every 10 items and at the end
        if (($exported % 10) -eq 0 -or $exported -eq $components.Count) {
            Write-Step ("exported {0}/{1}" -f $exported, $components.Count)
        }
    }

    $manifest = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $components.Count; $i++) {
        $c = $components.Item($i)
        $manifest.Add(("{0}`t{1}" -f $c.Name, $c.Type)) | Out-Null
    }

    $manifestPath = Join-Path $dir.FullName "_manifest.tsv"
    Write-Text $manifestPath (($manifest -join "`r`n") + "`r`n")

    $elapsed = (Get-Date) - $start
    Write-Step ("success: exported {0} components to {1} in {2:0.0}s" -f $exported, $OutDir, $elapsed.TotalSeconds)
    exit 0
}
catch {
    Write-Step ("FAILED: {0}" -f $_.Exception.Message)
    # Preserve original error details for logs/CI
    throw
}
finally {
    if ($null -ne $wb) { $wb.Close($false) | Out-Null }
    if ($null -ne $excel) { $excel.Quit() | Out-Null }

    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
