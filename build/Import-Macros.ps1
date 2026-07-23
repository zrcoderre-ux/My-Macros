<#
    Import-Macros.ps1
    Rebuilds My_Macros.dotm from the text modules in ..\src, and is aware of a
    running Word instance so it can hot-swap the template live.

    Called automatically by pull-extensions.ps1 after the repo updates, or run
    by hand anytime.

    Behavior:
      - Word CLOSED      -> rebuild the STARTUP .dotm on disk; next Word launch
        loads it. (A headless Word is started to do the build; it also loads the
        STARTUP template, so the add-in is unloaded first either way.)
      - Word OPEN, template loaded from STARTUP as a global add-in -> unload the
        add-in (releases the file lock), rebuild, then reload it, so the new
        macros are live without restarting Word.
      - Word OPEN with the .dotm open as a normal document -> skipped with a
        warning (close it and re-run), to avoid editing a file Word has locked.

    The build TARGET is the file Word actually loads: the copy in
    %AppData%\Microsoft\Word\STARTUP. The repo's src\ is the source. Override
    -Template if your Word Startup location is customized.

    Requires: Word installed, and Trust Center >
      "Trust access to the VBA project object model" ENABLED.
    Run under Windows PowerShell 5.1 (what pull-extensions.bat invokes), since
    it relies on Marshal::GetActiveObject to attach to a running Word.
#>

param(
    [string]$Template = (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\My_Macros.dotm"),
    [string]$Src      = (Join-Path $PSScriptRoot "..\src")
)

$ErrorActionPreference = "Stop"

$STD_MODULE   = 1
$CLASS_MODULE = 2
# (MSForm = 3, Document = 100)

$Template = [System.IO.Path]::GetFullPath($Template)
$Src      = (Resolve-Path $Src).Path
$leaf     = Split-Path $Template -Leaf

Write-Host "  [macros] rebuilding $leaf from src ..." -ForegroundColor Cyan

# --- Attach to a running Word, or start our own headless one ---------------
$word = $null; $startedWord = $false
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
    $word = New-Object -ComObject Word.Application
    $startedWord = $true
}
# Suppress alerts during the rebuild, but remember the prior value so the
# user's live Word instance gets it back on every exit path.
$prevDisplayAlerts = $word.DisplayAlerts
$word.DisplayAlerts = 0

# --- Release the lock on the STARTUP template -----------------------------
# Our own headless Word also loads STARTUP and locks the file, so this runs
# whether we attached to a live Word or started one.
$reinstallAddin = $false
$addin = $null

# Only the live instance can have it open as a normal document; bail if so.
if (-not $startedWord) {
    foreach ($d in $word.Documents) {
        if ($d.FullName -ieq $Template) {
            Write-Warning "  [macros] $leaf is open in Word as a document. Close it and re-run; skipped."
            try { $word.DisplayAlerts = $prevDisplayAlerts } catch {}
            return
        }
    }
}

# If it's loaded as a global add-in (STARTUP auto-loads it), unload to free it.
foreach ($a in $word.AddIns) {
    if ($a.Name -ieq $leaf) { $addin = $a; break }
}
if ($addin -and $addin.Installed) {
    $addin.Installed = $false
    $reinstallAddin = $true
    Write-Host "  [macros] unloaded global add-in to free the file." -ForegroundColor DarkGray
}

$doc = $null
$bootstrapped = $false
try {
    if (Test-Path $Template) {
        $doc = $word.Documents.Open($Template)
    } else {
        # First run / fresh clone: no .dotm in STARTUP yet. Create a blank one.
        Write-Host "  [macros] no $leaf in STARTUP; creating a fresh template." -ForegroundColor DarkGray
        $startupDir = Split-Path $Template -Parent
        if (-not (Test-Path $startupDir)) { New-Item -ItemType Directory -Path $startupDir -Force | Out-Null }
        $doc = $word.Documents.Add()
        $doc.SaveAs2($Template, 15)   # 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
        $bootstrapped = $true
    }
    $comps = $doc.VBProject.VBComponents

    # 1) Remove existing standard + class modules (keep ThisDocument + any form).
    $toRemove = @()
    foreach ($c in $comps) {
        if (($c.Type -eq $STD_MODULE) -or
            ($c.Type -eq $CLASS_MODULE -and $c.Name -ne "ThisDocument")) {
            $toRemove += $c.Name
        }
    }
    foreach ($n in $toRemove) { $comps.Remove($comps.Item($n)) }

    # 2) Re-import .bas and .cls (skip ThisDocument.cls; handled below).
    Get-ChildItem -Path $Src -Include *.bas, *.cls -File -Recurse |
        Where-Object { $_.Name -ne "ThisDocument.cls" } |
        ForEach-Object { $comps.Import($_.FullName) | Out-Null }

    # 3) Update ThisDocument code in place. A .cls export starts with a header
    #    block (VERSION / BEGIN..END / Attribute VB_*). Strip only the CONTIGUOUS
    #    header block at the top: stop at the first non-header line and keep
    #    everything from there on. (Matching header-ish lines ANYWHERE -- e.g. a
    #    mid-file "Attribute App.VB_VarHelpID = -1" emitted for WithEvents
    #    variables -- silently discarded all code above it.)
    $tdPath = Join-Path $Src "ThisDocument.cls"
    if (Test-Path $tdPath) {
        $lines = @(Get-Content -Path $tdPath)
        $firstBody = 0
        while ($firstBody -lt $lines.Count -and
               ($lines[$firstBody] -match '^\s*(VERSION |BEGIN|END\s*$|MultiUse|Attribute )' -or
                $lines[$firstBody] -match '^\s*$')) {
            $firstBody++
        }
        if ($firstBody -lt $lines.Count) {
            $body = ($lines[$firstBody..($lines.Count - 1)]) -join "`r`n"
        } else {
            $body = ""
        }
        $cm = $comps.Item("ThisDocument").CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($body)
    }

    # 4) Import the form ONLY if a real .frm + .frx pair exists in src.
    $frm = Join-Path $Src "frmSuggest.frm"
    $frx = Join-Path $Src "frmSuggest.frx"
    if ((Test-Path $frm) -and (Test-Path $frx)) {
        try { $comps.Remove($comps.Item("frmSuggest")) } catch {}
        $comps.Import($frm) | Out-Null
        Write-Host "  [macros] imported frmSuggest form." -ForegroundColor DarkGray
    }

    $doc.Save()
    $doc.Close()
    $doc = $null
    Write-Host "  [macros] rebuilt OK." -ForegroundColor Green
}
catch {
    Write-Warning "  [macros] rebuild failed: $_"
    Write-Warning "  [macros] Close Word, then run build\Import-Macros.ps1 by hand."
    if ($doc) { try { $doc.Close($false) } catch {} }
}
finally {
    if ($reinstallAddin -and $addin) {
        # Reload the existing add-in into the running Word so macros go live now.
        try { $addin.Installed = $true; Write-Host "  [macros] reloaded add-in (macros are live)." -ForegroundColor Green }
        catch { Write-Warning "  [macros] couldn't reload the add-in; restart Word to pick it up." }
    }
    elseif ($bootstrapped -and -not $startedWord) {
        # Brand-new template created while Word was open; load it live this session.
        try { $word.AddIns.Add($Template, $true) | Out-Null; Write-Host "  [macros] loaded new add-in (macros are live)." -ForegroundColor Green }
        catch { Write-Warning "  [macros] created the template; restart Word to load it." }
    }
    # Restore DisplayAlerts (kept suppressed through the add-in reload above
    # so that step stays silent too). Harmless on our own headless instance.
    try { $word.DisplayAlerts = $prevDisplayAlerts } catch {}
    if ($startedWord) {
        try { $word.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
