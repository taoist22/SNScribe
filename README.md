# SNScribe



https://github.com/user-attachments/assets/1558c891-adef-48bb-8eb3-32c893fe3c3d



**Write, format and cite in Word (.docx) documents on Supernote.**

SNScribe opens Word documents on a Supernote Manta or Nomad and lets you write in them — by
handwriting, the on-screen keyboard or a Bluetooth keyboard — then saves them back as Word
files that open cleanly in Word, Office 365 and LibreOffice. It is built for
students and writers: paper formats, citations from Zotero, quotes from your PDFs and EPUBs,
comments and tracked changes, handwritten margin notes and spell checking.

> **A drafting tool, not a finishing tool.** SNScribe is for writing, revising and formatting.
> Everything you set is saved as Word's own settings (fonts, styles, spacing, margins,
> headers, page numbers, the page breaks you insert), and every save is checked against what
> the screen shows. But SNScribe's pages are **screen pages**, not paper pages: Word decides
> where its own pages and lines end, so page count and line breaks will differ. Some things
> are also drawn only approximately on screen (hanging indents, raised superscript, colored
> highlights, tables and images) while being saved correctly. **Before you submit or print,
> open the document in Word, Office 365 or LibreOffice and check it there.**

Not affiliated with Microsoft. Word is a trademark of Microsoft Corporation.

---

## Install

1. Copy `SNScribe.snplg` (from the [releases](../../releases)) to the Supernote, for example into
   `MyStyle/`.
2. On the Supernote: **Settings → Apps → Plugins → Add Plugin**, and choose `SNScribe.snplg`.
3. Open any note or PDF/EPUB. The **SNScribe** button is on the toolbar (a plugin can't be started
   from the file manager).

The first time a feature needs it, the Supernote asks for permission: **files** (to open and save
documents) and **internet** (only for Zotero and DOI look-ups).

## Getting started

- **File → Open…** a `.docx`, or **File → New…** (blank) / **New from template…** (.docx or .dotx).
- **Tap** the page with the pen to place the caret and type. **Drag the pen** across words to
  select them; a drag inside one word selects just those letters. A **double tap** selects a word.
- **Swipe with a finger** to turn pages (left: next, right: previous). A finger never selects
  or types, so a hand resting on the screen does nothing.
- **Save** writes over the document. Before every save the previous version is kept
  (**File → Previous versions…**), and if the file was changed elsewhere since you opened it,
  SNScribe saves a copy instead of overwriting. **Save a copy** saves beside it.
- Unsaved work survives a crash or restart: SNScribe offers to **Restore** it next time.
- **◀ ▶** go to the previous and next page. The bottom line shows **Page N of M**; tap it (or
  **View → Go to page…**) to jump to a page. **Contents** lists the headings.
- **Pages** (tap the page count at the top) shows a card per page with its headings and how many
  comments, handwritten notes and tracked changes it has; **Only pages with notes** narrows the
  list. Card buttons (*New page after*, *Select page*, *Remove page break*) take a second tap.
- **Backspace** at the start of a page removes that page's page break.
- **View → Word count…** shows the word count, and can set a **target** for the body text
  (without a title page and the reference list); progress shows on the bottom line.
  **View** also changes the text size.

## Writing and formatting

| Where | What |
|---|---|
| Toolbar | **B I U**, **HL ▾** (highlight: yellow, green, blue, pink, red, turquoise, none), Undo / Redo |
| **Style ▾** | Body text, Heading 1–3, Quote, Title |
| **Para ▾** | Alignment, line spacing, space before/after, first-line or hanging indent, page breaks |
| **List ▾** | 1. 2. 3. · a. b. c. · A. B. C. · i. ii. iii. · outline (I. A. 1. a. i.) · bullets; **Indent / Outdent** — in a list, **Tab** and **Shift+Tab** do the same (lists continue as you type) |
| **Font ▾** | Strikethrough, superscript, subscript, font. With no text selected, Superscript/Subscript switch on for what you type next (type *H*, Subscript, *2*, Subscript, *O*) |
| **Size ▾** | Text size |
| **Edit ▾** | Cut, copy, paste, delete, **Find & replace**, **Spelling**, links, comments, citations, quotes, handwritten notes, **footnotes**, **pictures**, **tables**, accept/reject all changes |
| **File ▾** | Page setup (paper, margins, orientation), **Paper format (APA 7, MLA 9, Chicago)**, header/footer & page numbers, versions |

**Paper formats** set fonts, spacing, indents, margins, page numbers (with your last name for
MLA), the heading styles (e.g. APA: centered bold / left bold / left bold italic) and hanging
indents in the reference list — in one step you can undo.

Bluetooth keyboard: Ctrl + C / X / V / Z / Y / B / I / U, arrows, Shift + arrows to select,
Home / End, Delete. (The Cmd key on Mac-style keyboards doesn't work yet — use Ctrl.)

**Fonts.** SNScribe uses fonts from the Supernote's font folder (`MyStyle/Fonts`). For common Word
fonts that aren't on the device it can show a free look-alike with the same letter widths, so
lines and pages break as in Word: Calibri → **Carlito**, Cambria → **Caladea**, Arial →
**Liberation Sans**, Times New Roman → **Liberation Serif**, Courier New → **Liberation Mono**,
Georgia → **Gelasio**. Copy them (all four styles) into `MyStyle/Fonts`. The file keeps its fonts.

## Citations from Zotero

1. On zotero.org: **Settings → Security → Create new private key**, tick only *Allow library
   access* (read-only). Note the key and **your user ID** (the number shown there).
2. In SNScribe: **Edit → Cite from Zotero…**, enter the user ID and key — or **Load from file…**
   a `.txt` containing them — and **Save**. They are kept only on the Supernote.
3. Tap where the citation goes, search your library, choose **Parenthetical** or **Narrative**,
   add a page if you like, and **Insert**.

**Choosing the style:** **Citation style** buttons (APA 7, MLA 9, Chicago author-date, Chicago
notes) are at the top of **Edit → Cite from Zotero…** — also before Zotero is connected — and in
**Insert quote**, **Insert picture** and **Insert table**. It is one setting: SNScribe remembers it,
and citations, quotes and figure and table labels all follow it.

The citation goes into the sentence and the source into the **References** list (created on a
new page at the end the first time; *Works Cited* for MLA), in alphabetical order, formatted by
Zotero in APA 7, MLA 9 or Chicago author-date. Citing a source again never duplicates it.
Your library must sync to zotero.org; Wi-Fi is needed.

**Chicago (notes)** cites in footnotes instead: a full note the first time a source is cited,
a short note (*Author, Short Title, page.*) after that, and the source in a **Bibliography**.

## Quotes from your PDFs and EPUBs

1. In a PDF or EPUB, select a passage and tap **Quote → SNScribe** on the selection toolbar.
   Check the text (words broken across lines are rejoined) and the page, and **Keep quote**.
2. In your paper: **Edit → Insert quote…**, pick it. The first time for each file, say where its
   details come from — **Look up** its DOI (found on the first page when printed there), the
   **EPUB's own details**, **Find in Zotero**, or **Type the details** — then **Insert**.

Short quotes go into the sentence in quotation marks with (Author, Year, p. N); quotes of 40
words or more become a block quote. With Chicago (notes), a footnote follows the quotation
instead. The source joins the reference list. The files don't need to be in Zotero.

## Footnotes, pictures and tables

- **Footnotes:** numbers show raised in the text; **tap a number** to read, change or delete the
  footnote. **Edit → Footnote…** adds one at the caret. Italics in a note (a book title) stay
  where you don't retype the words.
- **Pictures** in a document show on the page (PNG, JPEG, GIF, BMP; Word's EMF/WMF drawings stay a
  placeholder). **Edit → Insert picture…** adds one after the paragraph you tapped — Small, Medium,
  Large or Full width (25–100% of the text width) — with an optional **figure label and title** in
  your citation style (APA: *Figure N*
  bold and the title in italics above; MLA and Chicago: a caption below). Later figures are
  renumbered; mentions like "see Figure 2" in your text are not.
  **Tap a picture** with the pen to change its size or delete it.
  To put space between a picture and the text below it, tap at the start of that text and press
  **Enter** (an empty line, which Word keeps too).
- **Tables** show as grids. **Tap a cell** to type in it, add a row above or below, delete a row,
  set the table's width (Full, 90%, 75%, 50%, centered), or delete the whole table. **Edit → Insert table…** adds one
  (rows × columns, width, header row, optional *Table N* label and title), ruled as APA tables are.
  Columns, merged cells and borders stay as Word made them. The cell lines on screen are a guide.
  A cell with links, pictures, fields or mixed formatting opens read-only (retyping it would lose
  them): edit that one in Word.

## Add to note

Opened from a note, SNScribe can put a passage into it: select the words, **Edit → Add to note…**,
then **Place in note** and tap where it goes on the note's page. It arrives as a Supernote text box
(searchable, editable; move or resize it with the lasso). Two switches, remembered:
**Add source line** puts a line under the passage — the citation of a quote in it (as your paper
cites it), else the document and the heading it's under — and **Stay in the note afterwards**
leaves you in the note to write about it (open SNScribe again to come back to the same place);
off, SNScribe comes straight back for the next passage. Opened from a PDF or EPUB, there's no
note to add to.

## Review: comments, tracked changes, handwritten notes

**Linked words:** the words a comment is on, the words a handwritten note is about, and quotes
you inserted have a **dashed underline** on screen (not in Word). **Tap them** to open the
comment, the note, or where the quote came from (file and page); **Type here** in that panel puts
the cursor where you tapped. Notes and quotes made before version 0.1.3 have no link.

- **Tracked changes** from Word show in the text (insertions underlined, deletions struck
  through) and as cards in the **margin panel**, each with ✓ accept / ✗ reject.
- **Comments** show as margin cards: tap to read the thread, reply or delete.
  **Edit → Comment…** adds one.
- **Handwritten notes** (**Edit → Handwritten note…**): write with the pen in the note box; your
  ink goes into the Word file as a picture in the right margin beside the words you chose (it
  shows and prints in Word), in the color you pick.
- The margin panel folds away with **› Fold panel**.

## Spelling

Misspellings are marked as you type (light grey, underlined; **View → Mark misspellings**
turns it off). **Edit → Spelling…** goes through them one by one: Change, Change all, Ignore,
Ignore all, **Add to dictionary**. English (US) is built in and works offline; other languages
are planned.

## Good to know

- Text boxes, Word drawings (EMF/WMF), tables inside tables, fields and content controls are
  kept exactly as they are and shown as placeholders. Endnotes show their numbers but can't be
  edited yet. Pictures can be resized but not moved or cropped.
- Formatting and paragraph-level tracked changes can't be reviewed yet; they stay as they are.
- Every save is checked: if the result would not match what the screen shows, nothing is
  written.
- A log for troubleshooting is written to `EXPORT/sn-docx-log-*.txt`.

## Privacy

Everything stays on the Supernote except: searches of **your** Zotero library
(api.zotero.org) and DOI look-ups (doi.org), and only when you use them. The Zotero key is
stored in the plugin's private storage.

## Develop

- `npx jest` — the editing, paging, citation and spelling logic.
- `cd android && ./gradlew :app:testDebugUnitTest` — the Kotlin reader, writer and spell checker
  against `fixtures/`. `DOCX_SAMPLES=/some/folder` also runs them against every `.docx` there.
- `./buildPlugin.sh`, then rename `build/outputs/SnDocx.snplg` to `SNScribe.snplg`
  (the build names the file after the internal key).

## Credits

- Icon: <a href="https://www.flaticon.com/free-icons/academic" title="academic icons">Academic icons created by nangicon - Flaticon</a> (`assets/academic-research.png`).
- English spelling dictionary: en_US from SCOWL (http://wordlist.sourceforge.net), by Kevin Atkinson and others,
  as distributed in LibreOffice's dictionaries; its copyright and license are in
  `android/app/src/main/resources/dict/LICENSE_en_US.txt`, which ships inside the plugin.

## License

[MIT](LICENSE)
