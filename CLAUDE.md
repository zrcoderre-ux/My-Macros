# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Workflow

- **Always open a pull request and merge it.** After committing work to the
  development branch, open a PR into `main` and merge it — do not leave changes
  sitting on the feature branch. The pull tool that rebuilds the `.dotm` only
  triggers when `main` changes, so work has to land on `main` to reach Word.
- Do **not** rewrite or reauthor commits that aren't Claude's own (e.g. the
  user's `zrcoderre@gmail.com` commits or GitHub's PR-merge commits), and do
  not force-push shared `main` history.

## Project model

- `src/` is the **source of truth** (`*.bas`, `*.cls`, `frmSuggest.frm/.frx`).
  Edit these. The `.dotm` is a build artifact rebuilt from `src/` on every pull
  and is **not** stored in the repo.
- The build target is the copy Word loads from its `STARTUP` folder, which Word
  auto-loads as a global add-in. `build/Import-Macros.ps1` rebuilds and
  hot-swaps it; `build/Export-Macros.ps1` regenerates `src/` after VBE edits.
- See `README.md` for the full build/setup flow.

## Code conventions

- This is Word VBA. Match the surrounding style: `Option Explicit`, explicit
  `#If VBA7` branches for Win32 declarations, and the existing comment density.
- Keep changes to a macro confined to that macro's module unless a shared helper
  is clearly the right call.
