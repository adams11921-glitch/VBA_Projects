# Session notes: how this project got here

Handoff from the first Claude Code session (cloud, October 4 2026,
https://claude.ai/code/session_01DXXotpdVc8pvbT96mrdAmg). A new session can read
this file to continue where that one stopped.

## The goal

Microsoft is ending Publisher. The owner (adams11921) wants a Word add-in that
looks and works like Publisher, mainly so volunteers at **Park Road Presbyterian
Church (PCA)** can keep making the weekly bulletin
(<https://parkroadpres.org/bulletins/>). The add-in will be sent to church members.

## What the real bulletin looks like (from Bulletin-10-4-26.pdf on the site)

* Folded, saddle-stapled booklet: **7 × 8.5" pages, two per side of landscape
  legal (8.5 × 14) paper**. The website PDF is 10 imposed sheets
  (e.g. 2|15, 14|3, …).
* Black heading bars with spaced text ("A Reflection Before the Service",
  "THIS WEEK at PARK ROAD", "UPCOMING THIS MONTH at PARK ROAD").
* Order of worship: bold items (The Prelude, The Welcome and Announcements, The
  Call to Worship: …, The Opening Hymns of Praise, …, The Sermon … — T. J.
  Campo). Hymns show the hymnal number at the right (#347). Songs not in the
  hymnal show the author and the full words, with the chorus indented. Scripture
  is in italics, followed by "Reader: The Word of the Lord / People: Thanks Be to
  God!" (People bold).
* Inside cover: reflection quotes, plus two copyright lines: "Scripture taken
  from the NEW AMERICAN STANDARD BIBLE" and "…Church Copyright License #337734".
* Announcements: "Event | Date | Time | Place", then details. A weekly schedule
  table (Day / Time / Event).
* Their street address was **not** found, so it is a placeholder in
  `DEF_ADDRESS` (`src/modPubBulletin.bas`).

## Key decisions

* **Booklet printing uses Word's built-in Book fold** (`PageSetup.BookFoldPrinting`).
  Pages are edited in reading order at finished size, and Word imposes them when
  printing, just as the owner remembered Publisher doing.
* The bulletin is **flowing text with custom paragraph styles**, not linked text
  boxes, so volunteers can add and remove items and everything reflows.
* The **Publisher tab** (free-form layout) uses page-anchored floating objects,
  with the section-1 header as the "master page" and dashed-line guides named
  `PubGuide_*`.
* Delivery: a `.dotm` installed into `%APPDATA%\Microsoft\Word\STARTUP` by
  `dist\Install.cmd`, which also unblocks the file.
* The owner's plan: the **bulletin editor uses Claude Code (with voice
  dictation)** to improve the add-in without the owner's help. Hence
  `CLAUDE.md`, `docs/LESSONS.md` (lessons accumulate each week) and the `/start`,
  `/wrap-up` and `/undo-last-change` commands. The owner intends to customize the
  start and close-out prompts in `.claude/commands/`.

## What was built (commits on branch `ccr-de12b3d9-x3bcla`)

1. Add-in: `src/*.bas` (4 modules), `ribbon/customUI14.xml` (Bulletin and
   Publisher tabs), `build/build_dotm.py`, `build/Add-Macros.ps1`, and
   `dist/Install.cmd` / `Uninstall.cmd`.
2. Booklet View (two pages side by side) and Check Pages (multiple-of-4 check
   that can add blank pages before the back cover). Print Booklet warns when the
   count is off.
3. Volunteer handout `docs/Bulletins-in-Word-Overview.pdf` (source
   `docs/overview.html`). Its pictures are **illustrations**, not real
   screenshots.
4. Claude Code workflow: `CLAUDE.md`, `docs/LESSONS.md`, `.claude/commands/*`,
   and `build/Update-AddIn.ps1` (one-step rebuild + compile check + install).

## Status: nothing has been run in Word yet

All the VBA was written without access to Word. **Next steps, on a Windows PC
with Word:**

1. Clone the repo and check out `ccr-de12b3d9-x3bcla` (or merge it to `main`).
   Open Claude Code in the `WordPublisherKit` folder.
2. In Word, turn on *Trust Center → Macro Settings → Trust access to the VBA
   project object model*.
3. Close Word and run `build\Update-AddIn.ps1`. Fix any compile errors it
   reports.
4. Open Word and try **Bulletin → New Bulletin → Legal paper**, the insert
   buttons, Check Pages, Print Booklet, and both PDFs. Fix what breaks, and write
   each fix in `docs/LESSONS.md`.
5. Fill in the real church address (`DEF_ADDRESS`) and confirm the other `DEF_…`
   values.
6. Optionally swap the handout illustrations for real screenshots.
7. Finish the owner's `/start` and `/wrap-up` prompts. Then set up the bulletin
   editor's PC and have them do a first session with the owner watching.

Things to double-check in Word, because they were written from memory of the
object model:
* `TextFrame.Next = …` (assigned without `Set`, as in Microsoft's docs example).
* Whether `PageSetup.PageWidth` reports 14" or 7" when Book fold is on.
  `GetPageSize` handles both.
* Placement of floating objects when pages are added or inserted (the anchor
  paragraph logic in `AddPageAtEnd` / `InsertPageAfterCurrent`).
* Some `imageMso` icon names (a wrong one only shows a blank icon).
* `ExportBookletPdf` switches the printer to "Microsoft Print to PDF". If that
  fails, it tells the user to pick it in the Print dialog.
