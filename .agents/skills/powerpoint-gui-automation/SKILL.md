---
name: powerpoint-gui-automation
description: Use on Windows with desktop Microsoft PowerPoint installed when a feature has NO COM/VBA surface at all (e.g. Insert > Zoom - Slide/Section/Summary Zoom - has no Shapes.AddZoom) and must be authored by driving the real GUI, or when you need to visually observe how a file renders/behaves in the actual desktop app rather than infer it from OOXML. Escalation path from powerpoint-fixture-authoring, not a replacement for it - try COM (and its ExecuteMso fallback) first.
---

# PowerPoint GUI Automation

Drives the real, visible desktop PowerPoint window: foreground control,
ribbon KeyTips over SendKeys, and UI Automation (`InvokePattern`/
`TogglePattern`) for anything keyboard and synthetic mouse can't reach.
Reconstructed 2026-07-21 from the session that authored the Zoom
(`pslz:sldZm` / `psez:sectionZm` / `psuz:summaryZm`) fixture, the first
feature found with zero COM surface. See [[powerpoint-com-cannot-author-zoom]]
memory for the ground-truth OOXML that came out of that session.

**Try COM first.** `powerpoint-fixture-authoring` covers ordinary COM
authoring and its `ExecuteMso`-on-a-selection fallback (for methods whose
enum arguments won't marshal from PowerShell, e.g. `MergeShapes`). Only reach
for full GUI automation when a feature genuinely has no COM/VBA surface -
verify that first by checking the real `Shapes`/`Presentation` object model
(e.g. the full `Add*` method list), not by assuming.

## Prerequisites

- Confirm the shell is in an **interactive console session**, not session 0
  or a disconnected one - a headless/service session renders nothing, so
  `CopyFromScreen` would silently return black or stale pixels:
  ```
  query session
  ```
  Look for `>console ... N  Active` (the `>` marks the current session). If
  it shows `services`/session 0 or a `Disc` state, GUI automation is not
  possible here - stop and say so rather than producing screenshots that
  look plausible but are empty.
- Desktop PowerPoint running with **a visible window** (not headless COM,
  not `WithWindow:=0`) - `Presentations.Open(path, msoFalse, msoFalse,
  msoTrue)` or an interactive `Presentations.Add()`.
- Clear stale recovery state first (same gotcha as COM authoring, see
  [[powerpoint-com-authoring-gotchas]]): delete
  `HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\{DocumentRecovery,StartupItems}`.
  A Document Recovery pane intercepts keyboard input before the ribbon does.
- Run PowerShell scripts through the **PowerShell (pwsh 7) tool** with the
  call operator (`& '...\script.ps1' ...`), same as the sibling skills - no
  `-ExecutionPolicy Bypass`, which trips the sandbox's weaken-security
  classifier. `uia-lib.ps1` needs `UIAutomationClient`, which requires an STA
  apartment; if you invoke via a raw `powershell.exe`/`pwsh` call instead of
  the tool's default host, add `-STA`.
- Keep script string literals **ASCII-only**. An em-dash inside a quoted
  string written to a `.ps1` file and run with `-File` corrupted the encoding
  and threw a spurious `TerminatorExpectedAtEndOfString` / `Missing closing
  '}'` parse error, far from the real problem.

## Workflow

1. **Sanity check.** `& '.agents\skills\powerpoint-gui-automation\scripts\foreground-and-shoot.ps1' -Shot .tmp\shot0.png`,
   then Read the PNG. Confirms the window foregrounds and you're actually
   looking at PowerPoint (a background/detached terminal will otherwise stay
   in front - plain `SetForegroundWindow` is blocked by Windows' foreground
   lock; the script uses the `AttachThreadInput` trick, see `ppt-window-lib.ps1`).
2. **Discover ribbon KeyTips.** Press Alt alone and screenshot - KeyTips are
   drawn as small overlay badges on every ribbon tab/button, so read them off
   the image rather than guessing:
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%' -Shot .tmp\shot-keytips.png
   ```
   Then Alt->tab-letter to reveal that tab's KeyTips, and so on down into any
   split-button dropdown.
3. **Drive the ribbon action + open dropdown/menu item.**
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%;N;Y' -UiaMenuItem 'Slide Zoom' `
       -DialogTitleLike 'Insert *Zoom*' -Dump -Shot .tmp\shot-dlg.png
   ```
   `-Dump` prints every CheckBox/Button/MenuItem/ListItem's REAL accessible
   `Name` in the dialog - **always run this before guessing a control's name
   from what's visible on screen.** Displayed text and accessible name can
   differ (a checkbox showing "2. Alpha 1" was actually named
   `"Slide 2 Alpha 1"` in the tree; guessing the visible label found nothing
   and the dialog's Insert button silently no-op'd on 0 selections - confirm
   via `-Dump`, not by assuming a click "probably worked").
4. **Toggle/submit using the real names from step 3's dump.**
   ```
   & drive-ribbon.ps1 -DialogTitleLike 'Insert *Zoom*' -Toggle 'Slide 2 Alpha 1' `
       -InvokeButton 'Insert' -Shot .tmp\shot-done.png
   ```
5. **Verify against the actual OOXML, not just the screenshot.** A
   screenshot proves the UI reacted; it does not prove the XML is correct or
   wired up (relationships, ids, namespaces).
   ```
   & save-and-extract.ps1 -NameLike 'my-fixture*' -DestDir .tmp\gui-extract
   ```
   then read the extracted `ppt/slides/slideN.xml` and `_rels` directly.
6. **Clean up / hygiene before committing anything GUI-authored:** a
   GUI-authored fixture's `docProps/core.xml` carries the interactively
   logged-in user's real name, and may carry add-in residue. Scrub before
   committing to a public repo. Snapshot pre-existing `POWERPNT` PIDs before
   you start (as in `powerpoint-fixture-authoring`'s reap pattern) so cleanup
   never kills a session you didn't spawn.

## Why keyboard/mouse alone don't work on ribbon popups (do not re-try these)

These were tried, in order, before UI Automation was found to be the
reliable path - re-reading this list before improvising a new approach will
save time:

1. Pixel-coordinate mouse click on a ribbon tab - registered but the ribbon
   didn't switch tabs.
2. Chorded `SendKeys` (`%N` as one call) - fires the OLD Alt+letter
   accelerator (Alt+N opened Header & Footer), not the KeyTip overlay.
   **KeyTips require Alt pressed-and-released, then each letter as its own
   sequential `SendKeys.SendWait` call** (see `Send-KeyTipSequence` in
   `ppt-window-lib.ps1`).
3. A submenu's KeyTip letter sent as part of the same sequential batch with a
   600ms gap, then 1200ms - never registered; the submenu's own KeyTips
   hadn't armed in time, or arrow-key navigation don't reach it at all.
4. Arrow keys (`{UP}`/`{ENTER}`) inside an open dropdown - no effect.
5. Synthetic mouse click (`SetCursorPos` + `mouse_event`) on the popup menu
   item's screen coordinates - the click registered (menu closed) but did
   **not** activate the item.
6. DPI/cursor-precision diagnosis of #5 (suspecting a coordinate offset) -
   ruled out; coordinates round-tripped exactly. The click itself is being
   swallowed, not mis-aimed.
7. "Hover then click" with an explicit move event and a dwell delay before
   the click - still no activation.

**Conclusion: Office ribbon dropdown popups (and likely their follow-on
dialogs) run on a separate input queue from both SendKeys and synthetic
mouse input** (consistent with UIPI integrity-level isolation). UI
Automation's `InvokePattern`/`TogglePattern` operate on the accessibility
tree directly, bypassing the input queue entirely, and is what actually
worked (`uia-lib.ps1`).

## Other gotchas

- **Full-virtual-screen screenshots and unscoped UIA desktop enumeration
  both leak content from the user's other open windows** (this happened:
  a Teams pane with confidential content was incidentally captured/
  enumerated before the scripts were scoped). `Save-WindowScreenshot` crops
  to the target window's `GetWindowRect`; `Find-UiaDialog` + the `Scope`
  param on every `uia-lib.ps1` helper restrict searches to one dialog by
  title. Always pass `-DialogTitleLike` / a scope - don't fall back to
  unscoped desktop search except as a last resort, and say so if you do.
- **Timing that mattered in this environment** (starting points, not gospel -
  re-tune if a step visibly hasn't settled in a screenshot): ~700ms after
  foregrounding before sending input; ~700-900ms between sequential KeyTip
  keystrokes (600ms was sometimes too short for a submenu to arm);
  ~1200-1500ms settle after a menu-opening Invoke before the next lookup;
  ~250-300ms between UIA toggles; ~1200-1300ms before the final screenshot.
- Loading `UIAutomationClient`/`UIAutomationTypes` flips the process
  DPI-aware mid-session, which can jump screenshot resolution (observed
  1280x800 -> 1920x1200). Harmless for UIA itself (coordinate-free by
  construction); only take a "before" screenshot ahead of dot-sourcing
  `uia-lib.ps1` if you need pixel-comparable before/after images.
- The sandbox's `Remove-Item` false-positive guard (blocks when the command
  text also contains regex-like substrings such as `r:` or `\w+` - easy to
  trip on when a nearby line quotes an XML rels attribute) applies here too;
  `save-and-extract.ps1` always extracts to a fresh directory rather than
  deleting an existing one.

## Scripts

- `scripts/ppt-window-lib.ps1` - dot-source library: `Get-PptHwnd`,
  `Set-PptForeground` (the `AttachThreadInput` foreground-lock workaround),
  `Save-WindowScreenshot` (window-rect-scoped capture), `Send-KeyTipSequence`.
- `scripts/uia-lib.ps1` - dot-source library: `Find-UiaDialog` (scope by
  window title), `Find-UiaElementByName`, `Get-UiaControlDump`,
  `Invoke-UiaElement`, `Set-UiaToggleOn`.
- `scripts/foreground-and-shoot.ps1` - standalone sanity check / "just look
  at current state" tool.
- `scripts/drive-ribbon.ps1` - the combined driver: KeyTips -> UIA menu
  invoke -> optional dump/toggle/button-invoke -> screenshot. See its
  comment-based help (`Get-Help ...\drive-ribbon.ps1 -Full`) for every
  parameter and worked examples.
- `scripts/save-and-extract.ps1` - saves the presentation via its already-
  running COM instance and extracts the package so you can read the real
  OOXML the GUI action produced, instead of trusting the screenshot alone.
