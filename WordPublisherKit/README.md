# Word Publisher Kit: Publisher-style tools and church bulletins in Word

Microsoft is retiring Publisher. This add-in puts two new tabs on the Word ribbon
so bulletins and other publications can be made in Word instead:

| Tab | For | What it does |
|---|---|---|
| **Bulletin** | Volunteers making the weekly bulletin | One click starts a folded bulletin (cover, reflection, order of worship, announcements). Buttons add hymns, Scripture readings, responsive readings, sermons, heading bars and announcements already formatted. It also prints the booklet in folding order. |
| **Publisher** | Flyers, posters, cards, brochures | Page-based layout like Publisher: text boxes and pictures that stay where you put them, linked text boxes, layout guides that don't print, a master page, align/arrange tools, and exact size & position. |

The bulletin layout follows the Park Road bulletin
(<https://parkroadpres.org/bulletins/>): 7 × 8.5" pages printed on landscape
legal paper and folded, with black heading bars, bold worship items, hymn numbers
aligned at the right, Reader/People responses, and announcement pages. A
5.5 × 8.5" version on letter paper is also included.

> **Not yet tested in Word.** The code was written and checked without access to
> Microsoft Word. Try it on one PC and make a test bulletin before sending it to the
> church. Expect a few small fixes after that first run.

---

## 1. One-time setup (you, on a Windows PC with Word)

The repository has the ribbon and the code as separate pieces. Put them together
once to make the add-in file you will send out.

1. Build the template with the ribbon (already done; it is in `dist/`):
   `python3 build/build_dotm.py`
2. Add the macros. Pick **either** option:
   * **Automatic.** In Word, go to *File → Options → Trust Center → Trust Center
     Settings → Macro Settings* and tick **Trust access to the VBA project object
     model**. Close Word. In PowerShell, run
     `powershell -ExecutionPolicy Bypass -File build\Add-Macros.ps1`.
     You can untick the setting afterwards.
   * **By hand.** Open `dist\WordPublisherKit.dotm` in Word and press **Alt+F11**.
     Then use *File → Import File…* on each of the four files in `src\`:
     `modPubCore.bas`, `modPubPages.bas`, `modPubObjects.bas`, `modPubBulletin.bas`.
     Use *Debug → Compile Project* to check for errors, save, and close Word.
3. Optional: set your church's details. In `src\modPubBulletin.bas`, edit the
   `DEF_…` lines at the top (church name, address, website, service times,
   copyright lines) **before** step 2. Each person can also change them later with
   **Bulletin → Church Settings**.
4. Test it. Run `dist\Install.cmd`, open Word, and click
   **Bulletin → New Bulletin → Legal paper**.

## 2. Sending it to church members

Send them three files from `dist\`, zipped together:

* `WordPublisherKit.dotm`, with the macros added in step 2
* `Install.cmd`
* `Uninstall.cmd`

Each person unzips the files, closes Word, and double-clicks **Install.cmd**.
This copies the add-in to Word's Startup folder, so the tabs appear every time
Word opens. It also unblocks the file so Windows doesn't block macros from email
or downloads.

Requirements: Windows with Word 2013 or newer (Microsoft 365 is fine). Mac Word
can load the ribbon, but the install script and some dialogs are Windows-only.

If someone's organisation blocks all macros, IT has to allow this one file. One
way is to sign it with a code-signing certificate.

## 3. Making a bulletin each week (volunteers)

1. **First time:** **Bulletin → New Bulletin → Legal paper**. Replace every
   `[bracketed]` placeholder. Click the grey picture box on the cover, then
   **Picture**, to add the cover image. Save the file, e.g. `Bulletin 2026-10-11.docx`.
2. **Every week after that:** open last week's bulletin and click
   **Start Next Week**. This saves a copy with the new date filled in and leaves
   last week's file unchanged. Then update the hymns, readings, sermon and
   announcements.
3. Add items where the cursor is with **Worship Item**, **Hymn**, **Scripture**,
   **Responsive Reading**, **Sermon**, **Heading Bar**, **Announcement**,
   **Weekly Schedule** and **New Page**. Remove an item by selecting its lines and
   pressing Delete. Everything reflows.
4. **Bulletin Styles** formats pasted or typed text, e.g. pasted hymn words →
   *Lyrics*, a chorus → *Chorus*, a congregation response → *People Line*.
5. Printing:
   * **Print Booklet**: prints double-sided on legal paper (flip on the
     **short edge**) in booklet order. Fold the stack and staple.
   * **PDF for Printing**: the same imposed pages as a PDF, for a copier or
     print shop. This is like the PDFs on the church website.
   * **PDF for Website**: one bulletin page per PDF page, easy to read on a
     phone.

Word pads the booklet with blank pages to a multiple of 4 automatically.

## 4. Publisher tab reference

| Group | Buttons |
|---|---|
| Publication | New Publication (letter, A4, flyer, tri-fold brochure, newsletter, postcard, business card, greeting card, poster, custom size), Page Setup |
| Pages | Add Page, Insert After, Delete Page (with everything on it), Pages Panel, Edit Master / Close Master (items on the master appear on every page) |
| Objects | Text Box, Picture, Picture Placeholder, Shapes, WordArt, Table |
| Linked Text | Link (select the story box, click Link, select an empty box, click Link), Break Link, Continue on Next Page, Check Overflow |
| Arrange | Bring to Front / Send to Back / Forward / Backward, Align (one object → to the page; several → to each other), Distribute, Wrap Text, Group / Ungroup, Selection Pane |
| Layout | Layout Guides (margins, columns, rows and gutter; they don't print), Show/Hide Guides, Grid, Snap to Grid, Size & Position (exact inches), Whole Page |
| Output | Print and Save as PDF (guides are hidden automatically) |

### How it works

* Each page ends with a page break, so every page starts with its own
  paragraph. Objects are anchored to their page's paragraph and positioned from
  the page corner, with "in front of text" wrapping. They stay where you put
  them, like in Publisher.
* The master page is the header of section 1. Layout guides are dashed lines drawn
  there, named `PubGuide_…`, and hidden while printing or exporting.
* Bulletins use normal flowing text with custom paragraph styles, plus Word's
  built-in *Book fold* page setup for printing in booklet order.

### Limits compared to Publisher

* No commercial print features (CMYK, spot colours, Pack and Go, crop marks).
* Word may still move objects if someone types in the body of a Publisher-tab
  publication. Keep content inside text boxes there.
* Very large Publisher-style documents can be slower in Word.
* Existing `.pub` files can't be opened. Re-create them with the tools above, or
  open them in Publisher before support ends and save as PDF or Word.

## Files

```
WordPublisherKit/
  src/                VBA modules (import into the .dotm)
    modPubCore.bas      ribbon dispatcher and shared helpers
    modPubPages.bas     new publications, pages, master page, guides, print/PDF
    modPubObjects.bas   text boxes, pictures, linked text, arrange/align/position
    modPubBulletin.bas  bulletin template, styles, worship/announcement inserts, booklet printing
  ribbon/customUI14.xml the Bulletin and Publisher ribbon tabs
  build/build_dotm.py   builds dist/WordPublisherKit.dotm with the ribbon
  build/Add-Macros.ps1  imports src/*.bas into the .dotm (Windows)
  dist/                 what you send out: WordPublisherKit.dotm, Install.cmd, Uninstall.cmd
```
