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
