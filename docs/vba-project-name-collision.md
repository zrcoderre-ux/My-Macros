# VBA Project Name Collision (`TemplateProject`)

> Filed for the record (August 2026). The build now renames this template's VBA
> project to `MyMacros`, so a fresh `Import-Macros.ps1` run fixes the collision
> permanently. Kept here because the symptom — the wrong macro running — points
> at everything except the real cause.

## The Problem

A document opens and the wrong macros run. Concretely, one or more of these:

- A keyboard shortcut (`Ctrl+Shift+V`, `Ctrl+Shift+T`, `Ctrl+Shift+C`,
  `Ctrl+Shift+H`, `Ctrl+Shift+D`, the spacebar and return wrap bindings) does
  nothing, or does something belonging to a different template.
- `Alt+F8` lists two entries with the same project name, or lists a macro twice.
- A macro that works fine in a blank document misbehaves in one specific
  document — typically a tentative draft that came out of a mail merge.

The tell is that it is document-specific. The add-in is loaded and healthy; a
particular document is pulling in a second project that outranks it.

## How to Confirm It

With the misbehaving document active, press `Alt+F11` to open the VBE and look
at the Project Explorer pane (`Ctrl+R` if it is hidden). A healthy session shows
this add-in's project and `Normal`. A collision shows a third project — the
document's attached template — and, before the fix below, two projects sharing
the name `TemplateProject`.

To see what the document is attached to, in that same window use the Immediate
pane (`Ctrl+G`):

```vba
? ActiveDocument.AttachedTemplate.FullName
```

A path ending in `Mail Merge Order Template (Automatic).dotm` (rather than
`Normal.dotm`) confirms it.

## The Actual Cause

Two things stack up.

First, **Word's default project name.** A template created through
`Documents.Add()` + `SaveAs2(..., 15)` — which is exactly how
`build/Import-Macros.ps1` bootstraps `My_Macros.dotm` — keeps Word's default
VBA project name, `TemplateProject`. The Mail Merge Order Template's project was
never renamed either, so it is also `TemplateProject`. Two loaded projects
sharing a name makes project-qualified macro and key-binding resolution
ambiguous.

Second, **attached-template precedence.** A document produced by a mail merge
inherits `AttachedTemplate` from the merge main document, so every saved draft
carries a pointer back at that `.dotm` in `word/_rels/settings.xml.rels`:

```
Type=".../attachedTemplate"
Target="file:///C:\Users\ZCoderre\Mail Merge Order Template (Automatic).dotm"
```

Word loads an attached template's VBA project alongside the document and gives
it precedence over a global add-in loaded from `STARTUP` when resolving a macro
name or a key binding. So when the names collide, the attached template wins and
this add-in loses.

## The Fix

Three pieces, all in this repo:

1. **The build renames the project.** `Invoke-TemplateRebuild` in
   `build/Import-Macros.ps1` sets `$doc.VBProject.Name = "MyMacros"` before
   importing any modules, so the name can never collide again. It is guarded by
   `try`/`catch` — a rename failure prints a note and the rebuild continues. The
   rename persists through the save at the end of the same function, and the
   message `VBA project renamed to MyMacros.` prints on the first run only.

   Nothing in `src/` hardcodes the old name, and `AutoExec` in
   `MacroKeyBindings` reapplies every key binding at each Word launch, so the
   rename orphans nothing.

2. **`AutoOpen` detaches on open.** `src/AutoOpen.bas` checks whether the
   document being opened is attached to a template whose name begins with
   `mail merge order template`; if so it clears `UpdateStylesOnOpen`, reattaches
   the document to `Normal`, and saves. The match is anchored to the start of the
   name so it fires on that one template and nothing else, and it runs before the
   folder loop because merged drafts land in `Downloads`, which the folder list
   does not cover.

3. **`DetachFromMergeTemplate` is the manual repair.** `Alt+F8` >
   `DetachFromMergeTemplate` does the same thing on demand for the active
   document, whatever template it is attached to. Use it for a document that
   `AutoOpen` did not catch. It clears `UpdateStylesOnOpen` first, so swapping to
   `Normal` cannot pull Normal's styles into the document on its next open.

## Verifying After a Rebuild

1. Run the pull shortcut, or by hand:
   `powershell -ExecutionPolicy Bypass -File "<repo>\build\Import-Macros.ps1"`.
   Expect `VBA project renamed to MyMacros.` on the first run only.
2. Open Word, press `Alt+F11`, and confirm the Project Explorer shows `MyMacros`
   and no second `TemplateProject`.
3. Open a merged draft saved before the fix. The VBE should no longer show the
   mail merge template's project, and every key binding should fire.

## Out of Scope Here

The leak at its source is in `Mail Merge Order Template (Automatic).dotm`, which
is binary and not version-controlled in this repo. Stopping the merge from
handing its `AttachedTemplate` to each output document, and clearing the stray
`Ctrl+S` binding stored in that template's `customizations.xml`, are by-hand
edits in the VBE. The pieces above are cleanup for drafts already on disk.
