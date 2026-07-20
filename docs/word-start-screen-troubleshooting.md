# Word Start Screen Troubleshooting Summary

> Filed for the record (July 2026). No evidence ties this to My_Macros.dotm —
> the safe-mode test below rules out everything in Word's STARTUP folder,
> including this template. Kept here so future troubleshooting starts with the
> full history. If the issue recurs, check Task Manager for a lingering
> WINWORD.EXE after quitting Word and note whether the parenthetical
> autocomplete popup was used that session (its Application.OnTime timer is the
> only macro feature that could theoretically keep a process alive; sessions
> tear down on document close as of PR #64).

## The Problem

Microsoft Word stopped displaying its Start screen (the panel showing templates and recent files) when the application launched. Instead, Word opened directly to a blank document every time. The change appeared suddenly with no known action on the user's part.

The setting that controls this behavior, "Show the Start screen when this application starts" (File > Options > General), was already checked. Unchecking and rechecking it made no difference. The checkbox would not take effect no matter how it was toggled, which indicated that something below the settings layer was overriding it or that Word was not honoring the setting at all.

## What Was Ruled Out

The troubleshooting worked through each plausible cause in order and eliminated the following:

- **Registry override (standard value).** The `DisableBootToOfficeStart` value under Word's `Options` key was present and already set to `0`, meaning the setting itself was correct. Word was being told to show the Start screen and ignoring it.
- **Add-ins.** Launching Word in safe mode still skipped the Start screen, which cleared any third-party add-in as the cause.
- **Shortcut switch.** The Word shortcut's Target field ended cleanly at `WINWORD.EXE` with no command-line switch appended, so nothing in the launch shortcut was bypassing the Start screen.
- **Group policy (per-user and machine-wide).** No `DisableBootToOfficeStart` policy value existed in either `HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Common\General` or the machine-wide equivalent. This ruled out an IT-pushed policy, and confirmed the court was not suppressing the Start screen centrally.
- **Office-wide update.** Excel opened to its Start screen normally. Because Word and Excel share the same policy path and update channel, an Office-wide change would have affected both. It did not, which isolated the problem to Word specifically.

## The Actual Cause

Word's `Options` registry key had gotten into a corrupted or stuck state. Although the Start screen value read `0` (correct), the key was not being honored on launch. This is a settings-level corruption isolated to Word.

The reason it was so difficult to fix, and why every "reopen and test" appeared to do nothing, was that background Office processes were keeping Word alive in memory. Each relaunch simply resurfaced the same running session rather than performing a true cold start, so none of the registry changes took effect and Word never regenerated its settings.

## The Solution

1. Close Word completely.
2. Open Task Manager (Ctrl+Shift+Esc) and end every Word-related process (`Microsoft Word` / `WINWORD.EXE`), including any running in the background. Closing Outlook and Teams also helps, since they can keep Office components alive.
3. Launch Word fresh.

With all background processes cleared, Word performed a genuine cold start, regenerated a clean `Options` key, and the Start screen returned.

If a clean cold start had still failed to restore the Start screen, the next step would have been an Office repair (Control Panel > Programs and Features > Microsoft Office > Change > Quick Repair, then Online Repair if needed). That was not necessary here.

## Likely Root Cause and Prevention

The corruption of the `Options` key generally cannot be traced to a single event, but the most common triggers are an Office update or reboot interrupting a write to the key, an ungraceful Word crash or hard close that failed to flush settings cleanly, or the lingering background processes themselves persisting stale state on exit. The last of these fits this case, since the fix only held after those processes were cleared.

Because the key has now been rebuilt clean, the issue should not recur. If it does return, that would point to something actively re-corrupting the key on each close, most likely a background process that is not exiting cleanly. In that event, making sure Word fully exits after each use, or running an Office repair, would be the durable fix.
