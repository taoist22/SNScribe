package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class DocxEditorTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-test").toFile()

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    private fun save(src: File, ops: List<Op>): Pair<File, DocxEditor.Saved> {
        val dest = File(work, "${src.nameWithoutExtension}-${System.nanoTime()}.docx")
        return dest to DocxEditor.save(src, ops, dest, File(work, "tmp"))
    }

    @Test
    fun formatsAcrossRunsAndObjects() {
        val src = File(fixtures, "probe-fixture.docx")
        val p2 = paragraphs(src)[2].text
        val start = p2.indexOf("make it bold")
        val end = p2.indexOf(", then") + ", then".length // spans the footnote reference (U+FFFC)
        val (dest, saved) = save(
            src,
            listOf(
                Op.Format(2, start, end, "h", true),
                Op.Format(1, 0, 5, "b", true), // "Plain"
                Op.Format(1, 47, 60, "b", false), // "a bold phrase" un-bolded
                Op.Style(5, "heading2"),
            ),
        )
        assertEquals(listOf("word/document.xml"), saved.changedParts)
        val ps = paragraphs(dest)
        val highlighted = ps[2].runs.filter { it.highlight }.joinToString("") { it.text }
        assertEquals(p2.substring(start, end), highlighted)
        assertTrue(ps[2].runs.single { it.obj == "note" }.highlight)
        assertEquals("Plain", ps[1].runs.first().text)
        assertTrue(ps[1].runs.first().bold)
        assertFalse(ps[1].runs.any { it.text.contains("bold phrase") && it.bold })
        assertEquals("heading", ps[5].kind)
        assertEquals(2, ps[5].level)
    }

    @Test
    fun normalRemovesTheHeadingStyle() {
        val src = File(fixtures, "probe-fixture.docx")
        val (dest, _) = save(src, listOf(Op.Style(0, "normal")))
        assertEquals("body", paragraphs(dest)[0].kind)
    }

    @Test
    fun textEditsKeepFormattingAndStructure() {
        val src = File(fixtures, "probe-fixture.docx")
        val before = paragraphs(src)
        val p1 = before[1].text // "Plain words come first …, then a bold phrase sits …"
        val bold = p1.indexOf("a bold phrase")
        val (dest, _) = save(
            src,
            listOf(
                Op.Text(1, 0, 5, "Simple"), // replace "Plain" (plain formatting)
                Op.Text(1, bold + 1 + 1, bold + 1 + 1 + "bold".length, "BRAVE"), // offsets after "Simple" (+1)
                Op.Text(1, 0, 0, ">> "), // insert at the very start
                Op.Text(3, before[3].text.length, before[3].text.length, " The end."), // append
                Op.Text(0, 0, before[0].text.length, ""), // delete a whole heading's text
            ),
        )
        val after = paragraphs(dest)
        assertTrue(after[1].text.startsWith(">> Simple words come first"))
        // Replacing inside the bold run keeps bold: the new word is bold.
        assertTrue(after[1].runs.any { it.text.contains("BRAVE") && it.bold })
        assertTrue(after[3].text.endsWith("differently formatted runs at once. The end."))
        assertEquals("", after[0].text)
        assertEquals("heading", after[0].kind) // the paragraph (and its style) stay
        assertEquals(before.size, after.size)
    }

    @Test
    fun refusesToDeleteANoteMarker() {
        val src = File(fixtures, "probe-fixture.docx")
        val p2 = paragraphs(src)[2].text
        val note = p2.indexOf(DocxReader.OBJECT)
        val dest = File(work, "refused.docx")
        val failure = runCatching { DocxEditor.save(src, listOf(Op.Text(2, note - 4, note + 2, "")), dest, File(work, "tmp")) }
        assertTrue(failure.isFailure)
        assertFalse(dest.exists()) // nothing unverified reaches the destination
        // Typing right after the marker is fine.
        val (ok, _) = save(src, listOf(Op.Text(2, note + 1, note + 1, "!")))
        assertTrue(paragraphs(ok)[2].text.contains("${DocxReader.OBJECT}!, then"))
    }

    @Test
    fun typesIntoAnEmptyParagraph() {
        val src = File(fixtures, "probe-fixture.docx")
        val empty = paragraphs(src).first { it.text.isEmpty() || it.runs.all { r -> r.obj != null } }
        val at = empty.text.length
        val (dest, _) = save(src, listOf(Op.Text(empty.index, at, at, "Caption\there")))
        assertTrue(paragraphs(dest).first { it.index == empty.index }.text.endsWith("Caption\there"))
    }

    /**
     * Stress (text): on every real document in DOCX_SAMPLES, in every paragraph with plain
     * editable text, replace the middle third, insert at the start and delete the last
     * character — skipping ranges that touch objects or fields, as the screen would.
     */
    @Test
    fun textEditsOnRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }
        for (f in files) {
            val ops = ArrayList<Op>()
            for (p in paragraphs(f)) {
                if (p.runs.any { it.locked || it.obj != null }) continue
                var t = p.text
                if (t.length < 6) continue
                val a = t.length / 3
                val b = 2 * t.length / 3
                ops.add(Op.Text(p.index, a, b, "«edited»"))
                t = t.substring(0, a) + "«edited»" + t.substring(b)
                ops.add(Op.Text(p.index, 0, 0, "▶ "))
                t = "▶ $t"
                ops.add(Op.Text(p.index, t.length - 1, t.length, ""))
            }
            val (dest, saved) = save(f, ops)
            println("${f.name}: ${ops.size} text ops verified, ${saved.changedParts}")
            assertTrue(dest.exists())
        }
    }

    @Test
    fun splitsAndJoinsParagraphs() {
        val src = File(fixtures, "probe-fixture.docx")
        val before = paragraphs(src)
        val t1 = before[1].text
        val cut = t1.indexOf("a bold phrase") + 2 // inside the bold run: "a " | "bold phrase…"
        val (dest, _) = save(
            src,
            listOf(
                Op.Split(1, cut),
                Op.Split(0, before[0].text.length), // Enter at the end of the heading
                Op.Join(4), // the old paragraph 2 (now 4) back onto the second half of paragraph 1 (now 3)
            ),
        )
        val after = paragraphs(dest)
        assertEquals(before.size + 1, after.size)
        assertEquals(before[0].text, after[0].text)
        assertEquals("", after[1].text)
        assertEquals("body", after[1].kind) // heading → its "next" style
        assertEquals(t1.substring(0, cut), after[2].text)
        assertEquals(t1.substring(cut) + before[2].text, after[3].text)
        assertTrue(after[3].runs.first().bold) // "bold phrase" kept its formatting
    }

    @Test
    fun refusesToJoinAcrossATable() {
        val src = File(fixtures, "probe-fixture.docx")
        // Paragraph 4 comes after the table.
        val failure = runCatching { DocxEditor.save(src, listOf(Op.Join(4)), File(work, "no.docx"), File(work, "tmp")) }
        assertTrue(failure.exceptionOrNull()?.message.orEmpty().contains("table"))
    }

    /**
     * Stress (paragraphs): on every real document, split every editable paragraph in the
     * middle (links included) — once keeping the splits, once joining each back at once.
     */
    @Test
    fun splitsAndJoinsRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }
        for (f in files) {
            val ps = paragraphs(f).filter { p -> p.text.length >= 4 && p.runs.none { it.locked } }
            // Highest index first, so earlier indexes stay valid.
            val splits = ps.sortedByDescending { it.index }.map { Op.Split(it.index, it.text.length / 2) }
            val roundTrip = ps.sortedByDescending { it.index }.flatMap { listOf(Op.Split(it.index, it.text.length / 2), Op.Join(it.index + 1)) }
            val (kept, _) = save(f, splits)
            val (back, _) = save(f, roundTrip)
            assertEquals(paragraphs(f).size + splits.size, paragraphs(kept).size)
            assertEquals(paragraphs(f).map { it.text }, paragraphs(back).map { it.text })
            println("${f.name}: ${splits.size} splits kept, ${roundTrip.size} split+join ops round-tripped")
        }
    }

    @Test
    fun makesNewNumberedAndBulletedLists() {
        val src = File(fixtures, "probe-fixture.docx")
        val (dest, saved) = save(
            src,
            listOf(
                Op.ListItem(1, "number", "n1"),
                Op.ListItem(2, "number", "n1"),
                Op.ListItem(3, "bullet", "b1"),
            ),
        )
        assertTrue(saved.changedParts.contains("word/numbering.xml"))
        val ps = paragraphs(dest)
        assertEquals("1.", ps[1].listLabel)
        assertEquals("2.", ps[2].listLabel)
        assertEquals(ps[1].numId, ps[2].numId)
        assertEquals("•", ps[3].listLabel)
        // Taking one out again.
        val (out, _) = DocxEditor.save(dest, listOf(Op.ListItem(2, "none", "")), File(work, "out.docx"), File(work, "tmp")).let { File(work, "out.docx") to it }
        assertEquals(null, paragraphs(out)[2].listLabel)
        assertEquals("1.", paragraphs(out)[1].listLabel)
    }

    @Test
    fun aListItemDropsTheParagraphsOwnIndent() {
        val src = File(fixtures, "probe-fixture.docx")
        val (dest, _) = save(src, listOf(Op.ParaProps(1, first = 720), Op.ListItem(1, "number", "n1")))
        val p = paragraphs(dest)[1]
        assertEquals("1.", p.listLabel)
        assertEquals(null, p.para.first)
    }

    @Test
    fun addsListDefinitionsToADocumentWithNone() {
        // The fixture without its numbering part, relationship and content type.
        val bare = File(work, "no-lists.docx")
        java.util.zip.ZipFile(File(fixtures, "probe-fixture.docx")).use { zip ->
            java.util.zip.ZipOutputStream(bare.outputStream()).use { out ->
                for (e in zip.entries()) {
                    if (e.name == "word/numbering.xml") continue
                    var bytes = zip.getInputStream(e).readBytes()
                    if (e.name == "word/_rels/document.xml.rels" || e.name == "[Content_Types].xml") {
                        bytes = String(bytes).replace(Regex("<(Relationship|Override)[^>]*numbering[^>]*/>"), "").toByteArray()
                    }
                    out.putNextEntry(java.util.zip.ZipEntry(e.name))
                    out.write(bytes)
                    out.closeEntry()
                }
            }
        }
        assertEquals(null, paragraphs(bare)[1].listLabel)
        val (dest, saved) = save(bare, listOf(Op.ListItem(1, "number", "n1"), Op.ListItem(2, "number", "n1")))
        assertTrue(saved.changedParts.containsAll(listOf("word/numbering.xml", "word/_rels/document.xml.rels", "[Content_Types].xml")))
        assertEquals(listOf("1.", "2."), paragraphs(dest).subList(1, 3).map { it.listLabel })
    }

    /** Stress (lists): number every third editable paragraph into one new list, and take existing items out. */
    @Test
    fun listEditsOnRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }
        for (f in files) {
            val ps = paragraphs(f)
            val ops = ps.filterIndexed { i, _ -> i % 3 == 0 }.map { p ->
                if (p.numId != null) Op.ListItem(p.index, "none", "") else Op.ListItem(p.index, "number", "n1")
            }
            val (dest, saved) = save(f, ops)
            val numbered = paragraphs(dest).filter { it.listLabel?.firstOrNull()?.isDigit() == true }.size
            println("${f.name}: ${ops.size} list ops, ${saved.changedParts}, $numbered numbered paragraphs")
            assertTrue(paragraphs(dest).filterIndexed { i, _ -> i % 3 == 0 }.all { p -> (p.numId == null) == (ps[p.index].numId != null) })
        }
    }

    @Test
    fun setsParagraphFormatting() {
        val src = File(fixtures, "probe-fixture.docx")
        val (dest, _) = save(
            src,
            listOf(
                Op.ParaProps(1, align = "justify", line = 480, lineRule = "auto", before = 240, after = 120, first = 720),
                Op.ParaProps(2, first = -720),
                Op.ParaProps(3, pageBreakBefore = true, align = "center"),
            ),
        )
        val ps = paragraphs(dest)
        assertEquals("justify", ps[1].align)
        assertEquals(DocxReader.ParaFmt(before = 240, after = 120, line = 480, lineRule = "auto", first = 720, pageBreakBefore = null), ps[1].para.copy(pageBreakBefore = null))
        assertEquals(-720, ps[2].para.first)
        assertEquals(720, ps[2].indentTwips) // the hanging indent brought its left indent
        assertEquals(true, ps[3].para.pageBreakBefore)
        assertEquals("center", ps[3].align)
        // Undoing the hanging indent takes the left indent it added back out.
        val (back, _) = DocxEditor.save(dest, listOf(Op.ParaProps(2, first = 0)), File(work, "back.docx"), File(work, "tmp")).let { File(work, "back.docx") to it }
        assertEquals(0, paragraphs(back)[2].indentTwips)
        assertEquals(null, paragraphs(back)[2].para.first)
    }

    /**
     * Stress: on every real document in DOCX_SAMPLES, bold the first word, highlight the
     * middle third and italicise the end of every paragraph, and make paragraph 1 a
     * Heading 2. save() re-reads and throws unless every paragraph's text is unchanged and
     * every other part is byte-identical. Skipped when DOCX_SAMPLES is unset.
     */
    @Test
    fun editsEveryParagraphOfRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }
        for (f in files) {
            val ps = paragraphs(f)
            val ops = ArrayList<Op>()
            for (p in ps) {
                val n = p.text.length
                if (n == 0) continue
                ops.add(Op.Format(p.index, 0, minOf(n, p.text.indexOf(' ').takeIf { it > 0 } ?: n), "b", true))
                if (n >= 3) ops.add(Op.Format(p.index, n / 3, 2 * n / 3, "h", true))
                if (n >= 2) ops.add(Op.Format(p.index, n / 2, n, "i", true))
            }
            ps.getOrNull(1)?.let { ops.add(Op.Style(it.index, "heading2")) }
            val t0 = System.nanoTime()
            val (dest, saved) = save(f, ops)
            val ms = (System.nanoTime() - t0) / 1_000_000
            val after = paragraphs(dest)
            val highlightedParas = after.count { p -> p.runs.any { it.highlight } }
            println("${f.name}: ${ops.size} ops, ${saved.changedParts}, $ms ms, highlighted paragraphs $highlightedParas/${after.count { it.text.length >= 3 }}")
            assertEquals(ps.map { it.text }, after.map { it.text })
        }
    }
}
