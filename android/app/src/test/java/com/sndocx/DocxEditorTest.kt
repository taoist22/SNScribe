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
