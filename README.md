# SNScribe (sn-docx)

SNScribe is a Word (.docx) document editor for Supernote: reading, editing, formatting, paper formats
(APA, MLA, Chicago), comments and tracked changes, handwritten margin notes, citations from
Zotero and quotes from PDFs/EPUBs, and spell checking. Plan: `../sn-docx-editor-PLAN.md`;
the feasibility evidence is in `../sn-docx-probe`.

Not affiliated with Microsoft. Word is a trademark of Microsoft Corporation.

**Milestone 1 (this):** open a `.docx` and read it a page at a time — headings, bold, italic,
underline, strikethrough, highlights, links, bullet and numbered lists, alignment and indents.
Tables and content controls show as labelled summaries; images and note references as symbols.
Contents lists the headings. Nothing is edited or saved yet.

Launch it from the toolbar of any note or PDF/EPUB (plugins cannot start from the file manager).

## Develop

- `npx jest` — paging logic.
- `cd android && ./gradlew :app:testDebugUnitTest` — the Kotlin reader against `fixtures/`.
  Set `DOCX_SAMPLES=/some/folder` to also dump every `.docx` there in a readable form
  (output in `android/app/build/test-results/`).
- `./buildPlugin.sh` → `build/outputs/SnDocx.snplg`. Log on the device: `EXPORT/sn-docx-log-*.txt`.

## Credits

- Icon: <a href="https://www.flaticon.com/free-icons/academic" title="academic icons">Academic icons created by nangicon - Flaticon</a> (`assets/academic-research.png`).
- English spelling dictionary: en_US from SCOWL (http://wordlist.sourceforge.net), by Kevin Atkinson and others,
  as distributed in LibreOffice's dictionaries; its copyright and licence are in
  `android/app/src/main/resources/dict/LICENSE_en_US.txt`, which ships inside the plugin.
