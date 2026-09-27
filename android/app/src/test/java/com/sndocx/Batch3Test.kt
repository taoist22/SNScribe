package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class Batch3Test {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-b3").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }
    private val src = File(fixtures, "probe-fixture.docx")

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    private fun save(ops: List<Op>, label: String): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"), paragraphs(src).map { it.text })
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    @Test
    fun highlightColoursStrikeAndScripts() {
        val out = save(
            listOf(
                Op.Format(1, 0, 5, "h", true, "green"),
                Op.Format(1, 6, 11, "h", true, "cyan"),
                Op.Format(3, 0, 4, "s", true),
                Op.Format(3, 5, 7, "sup", true),
                Op.Format(3, 8, 10, "sub", true),
                Op.Format(3, 8, 10, "sup", true), // superscript replaces subscript
            ),
            "b3-format",
        )
        val p1 = paragraphs(out)[1]
        assertEquals("green", p1.runs.first().highlightColor)
        assertTrue(p1.runs.any { it.highlightColor == "cyan" })
        val p3 = paragraphs(out)[3]
        assertTrue(p3.runs.first().strike)
        assertTrue(p3.runs.filter { it.superscript }.joinToString("") { it.text }.length == 4)
        assertFalse(p3.runs.any { it.subscript })
        // Un-highlighting writes "none".
        val cleared = save(listOf(Op.Format(1, 0, 5, "h", true, "red"), Op.Format(1, 0, 5, "h", false)), "b3-clear")
        assertFalse(paragraphs(cleared)[1].runs.first().highlight)
    }

    @Test
    fun runsReportTheirOwnFormattingApartFromTheStyle() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val typed = File(work, "typed.docx")
        DocxEditor.save(blank, listOf(Op.Text(0, 0, 0, "Heading text"), Op.Style(0, "heading1")), typed, File(work, "tmp"))
        val p = paragraphs(typed)[0]
        val run = p.runs.single()
        // The style's size and bold reach the paragraph, not the run's own formatting.
        assertEquals(null, run.ownSize)
        assertEquals(null, run.ownBold)
        assertTrue(run.bold)
        assertTrue(p.baseBold)
        assertEquals(32, p.baseSize)
    }

    @Test
    fun headingThreeAndQuote() {
        val out = save(listOf(Op.Style(2, "heading3"), Op.Style(3, "quote")), "b3-styles")
        val ps = paragraphs(out)
        assertEquals("heading", ps[2].kind)
        assertEquals(3, ps[2].level)
        assertEquals("body", ps[3].kind)
        assertTrue(ps[3].quote)
        // Back to Normal clears the quote.
        val back = save(listOf(Op.Style(3, "quote"), Op.Style(3, "normal")), "b3-unquote")
        assertFalse(paragraphs(back)[3].quote)
    }
}
