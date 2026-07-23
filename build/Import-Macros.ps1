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
      - Attached Word DIES mid-rebuild (RPC server unavailable, 0x800706BA --
        the user closed/crashed Word during the pull) -> automatically retry the
        build in a fresh private headless Word so the on-disk .dotm still gets
        rebuilt; the user just restarts Word to pick it up.

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

# ---------------------------------------------------------------------------
# True when an error is a dead/disconnected COM instance -- the Word process we
# were driving terminated (RPC server unavailable / call failed / disconnected).
# Matched by HRESULT (locale-independent) with a message-text backstop.
# ---------------------------------------------------------------------------
function Test-DeadCom {
    param($errRec)
    $ex = $errRec.Exception
    while ($ex) {
        if ($ex -is [Runtime.InteropServices.COMException]) {
            # 0x800706BA RPC_S_SERVER_UNAVAILABLE, 0x800706BE RPC_S_CALL_FAILED,
            # 0x80010108 RPC_E_DISCONNECTED.
            if (@(-2147023174, -2147023170, -2147417848) -contains $ex.HResult) { return $true }
        }
        $ex = $ex.InnerException
    }
    return ("$errRec" -match 'RPC server is unavailable|remote procedure call failed|disconnected from its clients|0x800706BA|0x800706BE|0x80010108')
}

# ---------------------------------------------------------------------------
# Unload the STARTUP template if $w has it loaded as a global add-in, to release
# the file lock. Both an attached live Word and our own headless Word auto-load
# STARTUP, so this is needed either way. Returns the AddIn object (so a live
# instance can reload it afterward), or $null if it wasn't loaded.
# ---------------------------------------------------------------------------
function Unload-StartupAddin {
    param($w)
    $a = $null
    foreach ($x in $w.AddIns) { if ($x.Name -ieq $leaf) { $a = $x; break } }
    if ($a -and $a.Installed) {
        $a.Installed = $false
        Write-Host "  [macros] unloaded global add-in to free the file." -ForegroundColor DarkGray
    }
    return $a
}

# ---------------------------------------------------------------------------
# The actual rebuild against Word instance $w: open (or create) the template,
# swap every module for the src copies, save, close. Assumes the file lock is
# already released (call Unload-StartupAddin first). Returns $true if it had to
# bootstrap a brand-new template. Closes the doc and rethrows on any failure.
# ---------------------------------------------------------------------------
function Invoke-TemplateRebuild {
    param($w)

    $didBootstrap = $false
    $doc = $null
    try {
        if (Test-Path $Template) {
            $doc = $w.Documents.Open($Template)
        } else {
            # First run / fresh clone: no .dotm in STARTUP yet. Create a blank one.
            Write-Host "  [macros] no $leaf in STARTUP; creating a fresh template." -ForegroundColor DarkGray
            $startupDir = Split-Path $Template -Parent
            if (-not (Test-Path $startupDir)) { New-Item -ItemType Directory -Path $startupDir -Force | Out-Null }
            $doc = $w.Documents.Add()
            $doc.SaveAs2($Template, 15)   # 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
            $didBootstrap = $true
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
    }
    catch {
        if ($doc) { try { $doc.Close($false) } catch {} }
        throw
    }

    return $didBootstrap
}

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

# Only the live instance can have the template open as a normal document; bail if so.
if (-not $startedWord) {
    foreach ($d in $word.Documents) {
        if ($d.FullName -ieq $Template) {
            Write-Warning "  [macros] $leaf is open in Word as a document. Close it and re-run; skipped."
            try { $word.DisplayAlerts = $prevDisplayAlerts } catch {}
            return
        }
    }
}

$addin          = $null
$reinstallAddin = $false
$bootstrapped   = $false
$builtOK        = $false

try {
    $addin = Unload-StartupAddin $word
    # A live instance whose add-in we unloaded gets it reloaded in finally, even
    # if the build then fails -- set the flag now so that recovery still happens.
    if ($addin -and -not $startedWord) { $reinstallAddin = $true }

    $bootstrapped = Invoke-TemplateRebuild $word
    $builtOK = $true
    Write-Host "  [macros] rebuilt OK." -ForegroundColor Green
}
catch {
    $firstErr = $_
    if ((-not $startedWord) -and (Test-DeadCom $firstErr)) {
        # The Word we were driving died mid-rebuild (typically the user closed or
        # crashed Word during the pull). Its file lock and add-in state died with
        # it, so a fresh private Word can rebuild the on-disk .dotm cleanly.
        Write-Warning "  [macros] the running Word became unavailable mid-rebuild; retrying in a private headless Word ..."
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
        $word = $null; $addin = $null; $reinstallAddin = $false
        try {
            $word = New-Object -ComObject Word.Application
            $startedWord = $true
            $prevDisplayAlerts = $word.DisplayAlerts
            $word.DisplayAlerts = 0
            $null = Unload-StartupAddin $word            # free the file on the new instance
            $bootstrapped = Invoke-TemplateRebuild $word
            $builtOK = $true
            Write-Host "  [macros] rebuilt OK (headless). Restart Word to load the new macros." -ForegroundColor Green
        }
        catch {
            Write-Warning "  [macros] headless rebuild also failed: $_"
        }
    }
    else {
        Write-Warning "  [macros] rebuild failed: $firstErr"
    }
    if (-not $builtOK) {
        Write-Warning "  [macros] Close Word (end any WINWORD.EXE in Task Manager), then run build\Import-Macros.ps1 by hand."
    }
}
finally {
    if ($reinstallAddin -and $addin -and $word) {
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
    if ($word) { try { $word.DisplayAlerts = $prevDisplayAlerts } catch {} }
    if ($startedWord -and $word) {
        try { $word.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
