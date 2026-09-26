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
