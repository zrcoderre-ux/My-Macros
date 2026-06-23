<#
    Export-Macros.ps1
    Exports every VBA component from My_Macros.dotm to ..\src as text.

    Run this when you've edited macros directly in Word and want the
    repo's text source to catch up (the reverse of Import-Macros.ps1).

    Same Trust Center requirement as the importer.
    NOTE: this is the ONLY reliable way to get a real, importable
    frmSuggest.frm + frmSuggest.frx pair, since the VBE writes them itself.
#>

param(
    [string]$Template = (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\My_Macros.dotm"),
    [string]$Src      = (Join-Path $PSScriptRoot "..\src")
)

$ErrorActionPreference = "Stop"
$Template = (Resolve-Path $Template).Path
if (-not (Test-Path $Src)) { New-Item -ItemType Directory -Path $Src | Out-Null }
$Src = (Resolve-Path $Src).Path

$ext = @{ 1 = ".bas"; 2 = ".cls"; 3 = ".frm"; 100 = ".cls" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Template)
    foreach ($c in $doc.VBProject.VBComponents) {
        $e = $ext[[int]$c.Type]
        if (-not $e) { continue }
        $out = Join-Path $Src ($c.Name + $e)
        $c.Export($out)              # writes .frx alongside automatically for forms
        Write-Host "  exported $($c.Name)$e"
    }
    $doc.Close($false)
    Write-Host "Done. src refreshed from $Template." -ForegroundColor Green
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
