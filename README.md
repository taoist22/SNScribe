# SNScribe

**Write, format and cite in Word (.docx) documents on Supernote.**

SNScribe opens Word documents on a Supernote Manta or Nomad and lets you write in them — by
handwriting, the on-screen keyboard or a Bluetooth keyboard — then saves them back as Word
files that open cleanly in Word, Office 365 and LibreOffice. It is built for
students and writers: paper formats, citations from Zotero, quotes from your PDFs and EPUBs,
comments and tracked changes, handwritten margin notes and spell checking.

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
- **Tap** the page to place the caret and type. **Drag the pen** across words to select them;
  a drag inside one word selects just those letters. A **double tap** selects a word.
- **Save** writes over the document. Before every save the previous version is kept
  (**File → Previous versions…**), and if the file was changed elsewhere since you opened it,
  SNScribe saves a copy instead of overwriting. **Save a copy** saves beside it.
- Unsaved work survives a crash or restart: SNScribe offers to **Restore** it next time.
- **◀ ▶** turn pages; **Contents** lists the headings; the page count opens **Pages**
  (thumbnails, *New page after*, *Select page*). **View** changes the text size.

## Writing and formatting

| Where | What |
|---|---|
| Toolbar | **B I U**, **HL ▾** (highlight: yellow, green, blue, pink, red, turquoise, none), Undo / Redo |
| **Style ▾** | Body text, Heading 1–3, Quote, Title |
| **Para ▾** | Alignment, line spacing, space before/after, first-line or hanging indent, page breaks |
| **List ▾** | Numbered, bulleted (lists continue as you type) |
| **Font ▾** | Strikethrough, superscript, subscript, font. With no text selected, Superscript/Subscript switch on for what you type next (type *H*, Subscript, *2*, Subscript, *O*) |
| **Size ▾** | Text size |
| **Edit ▾** | Cut, copy, paste, delete, **Find & replace**, **Spelling**, links, comments, citations, quotes, handwritten notes, accept/reject all changes |
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

The citation goes into the sentence and the source into the **References** list (created on a
new page at the end the first time; *Works Cited* for MLA), in alphabetical order, formatted by
Zotero in APA 7, MLA 9 or Chicago author-date. Citing a source again never duplicates it.
Your library must sync to zotero.org; Wi-Fi is needed.

## Quotes from your PDFs and EPUBs

1. In a PDF or EPUB, select a passage and tap **Quote → SNScribe** on the selection toolbar.
   Check the text (words broken across lines are rejoined) and the page, and **Keep quote**.
2. In your paper: **Edit → Insert quote…**, pick it. The first time for each file, say where its
   details come from — **Look up** its DOI (found on the first page when printed there), the
   **EPUB's own details**, **Find in Zotero**, or **Type the details** — then **Insert**.

Short quotes go into the sentence in quotation marks with (Author, Year, p. N); quotes of 40
words or more become a block quote. The source joins the reference list. The files don't need
to be in Zotero.

## Review: comments, tracked changes, handwritten notes

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

- Tables, pictures and text boxes in a document are kept exactly as they are and shown as
  placeholders; they can't be edited yet. Footnotes, fields and content controls are kept too.
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
