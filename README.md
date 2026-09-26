# DOCX (sn-docx)

A Word document reader for Supernote, growing into an editor. Plan: `../sn-docx-editor-PLAN.md`;
the feasibility evidence is in `../sn-docx-probe`.

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
