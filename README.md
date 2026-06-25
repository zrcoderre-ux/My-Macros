# My_Macros

VBA source for the Word macro template `My_Macros.dotm`, version-controlled as
text and wired into the one-click GitHub auto-pull tool so editing a macro is as
seamless as editing a Chrome extension.

## The model

The macros do not live in editable form inside the `.dotm` (they sit in a binary
`vbaProject.bin`). So the **text in `src/` is the source of truth**, and the
`.dotm` is a **build artifact** rebuilt from `src/` on every pull. The build
target is the copy Word actually loads, in `%AppData%\Microsoft\Word\STARTUP`,
which Word auto-loads as a global add-in at launch. No copy of the `.dotm` lives
in the repo.

This mirrors the Chrome flow. The pulled files are not what the host app runs
until something reloads them: Chrome reloads the extension; Word reloads the
template. `build/Import-Macros.ps1` does that Word-side reload.

## Layout

    src/                     source of truth (edit these)
      *.bas  *.cls           standard, class, and document modules
      frmSuggest.frm/.frx    UserForm (after you seed it once; see setup)
    build/
      Import-Macros.ps1      rebuild .dotm from src + hot-swap into running Word
      Export-Macros.ps1      regenerate src from the .dotm (run after VBE edits)
    My_Macros.dotm           NOT in repo; built into Word's STARTUP folder

## One-time setup

1. Enable **Trust access to the VBA project object model** (Word > Options >
   Trust Center > Trust Center Settings > Macro Settings).
2. Seed the form so it lives in the repo and round-trips. From a machine with
   the real template loaded, run `build/Export-Macros.ps1` once, then commit the
   resulting `src/frmSuggest.frm` and `src/frmSuggest.frx`.
3. In `pull-extensions.ps1`, set the My_Macros entry's `Url` and `Path`.

That's it. Because the build target is the STARTUP folder, Word auto-loads the
template as a global add-in. There is no manual "Templates and Add-ins" step,
and the first build creates the STARTUP `.dotm` from scratch if none exists.
If your Word Startup location is customized (Word > Options > Advanced > File
Locations > Startup), pass that path to the importer via `-Template`.

## Day-to-day

Edit `.bas` / `.cls` in `src/` (Claude Code or any editor), merge to `main`,
then click your existing pull shortcut. The tool pulls every repo; when this one
changes it rebuilds the `.dotm` and, if Word is open with the add-in loaded,
unloads/rebuilds/reloads it so the macros are live with no restart. If Word is
closed, the next launch loads the rebuilt template.

If you instead edit a macro inside Word's VBE, run `build/Export-Macros.ps1`
first to pull that change back into `src/`, then commit.

## Notes

- The importer triggers only when `main` actually changes, so your edits must
  land on `main` (your auto-merge-to-main workflow does this).
- If Word has the `.dotm` open as a normal document (not as the add-in), the
  rebuild is skipped with a warning; close it and re-run.
- If a rebuild ever fails, nothing is lost: fix the issue, close Word, and run
  `build/Import-Macros.ps1` by hand.

## Troubleshooting: a change merged but Word still runs the old macro

If you merged a change to `main` but Word still behaves the old way, the
STARTUP `.dotm` didn't actually get rebuilt. Work through these in order:

1. **Confirm the loaded build is stale.** Check when the template Word loads was
   last built:

   ```powershell
   Get-Item "$env:APPDATA\Microsoft\Word\STARTUP\My_Macros.dotm" |
     Select-Object FullName, LastWriteTime
   ```

   If `LastWriteTime` predates your change, Word is running an old build.

2. **Make sure your local clone actually pulled.** The importer builds from the
   local `src/`, so a failed `git pull` leaves it stale. If `git pull` errors
   with *"Please specify which branch you want to merge with"*, the local `main`
   has no upstream; fix it once:

   ```powershell
   git branch --set-upstream-to=origin/main main
   git pull origin main
   ```

3. **Run the importer by hand with the execution policy bypassed.** Running the
   script directly (not via the pull shortcut) hits PowerShell's default policy
   and fails with *"not digitally signed"*. Invoke it like this (this also forces
   Windows PowerShell 5.1, which the script requires):

   ```powershell
   powershell -ExecutionPolicy Bypass -File "<repo>\build\Import-Macros.ps1"
   ```

4. **If it fails with `RPC server is unavailable (0x800706BA)`**, the script
   attached to a live or zombie Word that died mid-rebuild. Give it a clean slate
   so it starts its own headless Word, then re-run step 3:

   ```powershell
   Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
   Get-Process WINWORD -ErrorAction SilentlyContinue   # should print nothing
   ```

5. **If it fails with a VBProject/"hidden module" or access error**, enable
   Word > Options > Trust Center > Trust Center Settings > Macro Settings >
   **Trust access to the VBA project object model**, then re-run step 3.

A successful run prints `rebuilt OK` and the STARTUP `.dotm` `LastWriteTime`
updates to now. Launch Word and confirm the new behavior. To verify which code
is actually loaded, press `Alt+F11` in Word and inspect the module directly.
