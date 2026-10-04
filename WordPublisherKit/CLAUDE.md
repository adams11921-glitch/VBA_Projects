# Word Publisher Kit: instructions for Claude

@docs/LESSONS.md

## Who you are working with

The person in this session puts together the weekly bulletin for Park Road
Presbyterian Church. They are **not a programmer**. They often talk to you
through voice dictation, so expect transcription mistakes ("hymn" might come
through as "him", "add-in" as "adding"). When a request could mean two things,
ask one short question before changing anything.

* Use plain language. Don't show code unless they ask for it. Say what will
  change in the bulletin or on the ribbon, not which function changed.
* Make **one small change at a time**. Rebuild and reinstall the add-in, then
  tell them exactly what to click in Word to try it.
* Before a large change (new feature, restructuring), describe the plan in 2–3
  sentences and wait for "yes".
* Never delete or overwrite their bulletin files (`*.docx`) or anything outside
  this folder.
* If something breaks, `/undo-last-change` puts the previous version back.
  Remind them of this when a change is risky.

## What this project is

A Word add-in (a macro-enabled template, `.dotm`) that replaces Microsoft
Publisher. It adds a **Bulletin** tab (weekly folded-booklet bulletins) and a
**Publisher** tab (free-form layout) to Word's ribbon. See `README.md` for the
features and `docs/Bulletins-in-Word-Overview.pdf` for the volunteer handout.

| Path | What it is |
|---|---|
| `src/modPubCore.bas` | Ribbon dispatcher `RunCommand` (every button lands here) and shared helpers |
| `src/modPubBulletin.bas` | Bulletin: New Bulletin layout, styles, insert buttons, dates, booklet printing, church settings (`DEF_…` constants at the top) |
| `src/modPubPages.bas` | Publications, pages, master page, layout guides, print/PDF |
| `src/modPubObjects.bas` | Text boxes, pictures, linked text, arrange/align, size & position |
| `ribbon/customUI14.xml` | Ribbon buttons |
| `dist/WordPublisherKit.dotm` | The built add-in that gets installed and sent to others |
| `build/Update-AddIn.ps1` | **Run after every change.** Puts the ribbon and code into the .dotm and installs it |
| `docs/LESSONS.md` | Lessons learned. Read it at the start, add to it at `/wrap-up` |

## How to make a change

1. Edit `src/*.bas` and/or `ribbon/customUI14.xml`.
2. To add a button: give it a unique `id` in the XML with
   `onAction="Pub_OnAction"`, then add `Case "<id>": <SubName>` in `RunCommand`
   (`modPubCore.bas`). A button that appears on both tabs uses a different `id`,
   `onAction="Pub_OnTagAction"`, and `tag="<command id>"`.
3. Ask them to **close Word**, then run
   `powershell -NoProfile -ExecutionPolicy Bypass -File build\Update-AddIn.ps1`.
   This updates the ribbon, imports the modules, compiles, and installs into
   Word's STARTUP folder. If it reports a compile error, fix it and run it again
   before handing back.
4. Optional smoke test without clicking: after installing, you can drive Word
   from PowerShell (`$w = New-Object -ComObject Word.Application;
   $w.Run("NewBulletin", "Legal")`, export a PDF with
   `$w.ActiveDocument.ExportAsFixedFormat(...)`, render it, and look at it). Avoid
   macros that open InputBox/MsgBox, which block an invisible Word.
5. Tell them what to click to try it. Once they're happy, commit with a plain
   message (`git commit -am "Hymn button: add verse numbers"`).

## VBA rules for this project (learned the hard way)

* Keep `.bas` files with **CRLF** line endings, or the VBA editor imports them as
  garbage. `.gitattributes` keeps them as-is.
* A single compile error disables **every** button. If you're unsure a Word
  property exists in all versions, call it through a late-bound `Object`
  variable (see `ToggleGrid`, `IsOverflowing`).
* Inside a Function, don't read the function's own name as a value (it can
  recurse). Use a local variable and assign it at the end.
* Errors: call `PubFail "message"` for problems the user should see. `RunCommand`
  shows them in a friendly message box.
* Bulletin text is normal flowing text formatted with the styles in
  `EnsureBulletinStyles`. Styles are only created if missing, so a church's own
  tweaks to a style survive.
* Booklet printing relies on Word's `PageSetup.BookFoldPrinting`. Pages stay in
  reading order on screen, and Word imposes them when printing. Page count should
  be a multiple of 4 (`CheckBookletPages`).
* Ribbon `id`s must be unique. Built-in `imageMso` names that don't exist just
  show a blank icon.

## Sharing an update with other volunteers

After a change has been tested, the file to send is `dist\WordPublisherKit.dotm`
(with `Install.cmd` and `Uninstall.cmd`). Recipients close Word and run
`Install.cmd` again. Update `README.md` and the handout
(`docs/overview.html` → PDF) when a feature visibly changes.
