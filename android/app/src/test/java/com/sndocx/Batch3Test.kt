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
    fun styleLooksComeFromTheFile() {
        val blank = File(work, "blank-looks.docx").also { DocxBlank.write(it) }
        val looks = DocxReader.read(blank).looks
        assertEquals(32, looks.getValue("heading1").size)
        assertTrue(looks.getValue("quote").italic)
        assertEquals(720, looks.getValue("quote").indent)
        // CT's IDS 105 template (DOCX_SAMPLES): Heading 1 is 12 pt bold centred; Heading 3
        // is body size, bold, centred and indented — and no paragraph uses it yet.
        val dir = System.getProperty("docx.samples").orEmpty()
        val ids = File(dir, "IDS 105 Project Template.docx")
        if (dir.isEmpty() || !ids.isFile) return
        val l = DocxReader.read(ids).looks
        assertEquals(24, l.getValue("heading1").size)
        assertEquals("center", l.getValue("heading1").align)
        assertEquals(22, l.getValue("heading3").size)
        assertEquals("center", l.getValue("heading3").align)
        assertEquals(360, l.getValue("heading3").indent)
    }

    @Test
    fun apaRedefinesHeadingStyles() {
        val apa = listOf(
            DocxEditor.StyleDef("heading1", "Times New Roman", 24, bold = true, italic = false, align = "center", line = 480),
            DocxEditor.StyleDef("heading2", "Times New Roman", 24, bold = true, italic = false, align = "left", line = 480),
            DocxEditor.StyleDef("heading3", "Times New Roman", 24, bold = true, italic = true, align = "left", line = 480),
        )
        val dir = System.getProperty("docx.samples").orEmpty()
        val ids = File(dir, "IDS 105 Project Template.docx")
        val sources = listOfNotNull(File(work, "blank-apa.docx").also { DocxBlank.write(it) }, ids.takeIf { dir.isNotEmpty() && it.isFile })
        for ((n, src) in sources.withIndex()) {
            val dest = File(work, "apa-$n.docx")
            DocxEditor.save(src, listOf(Op.StyleDefs(apa)), dest, File(work, "tmp"), paragraphs(src).map { it.text })
            dest.copyTo(File(keep, "b3-apa-styles-$n.docx"), overwrite = true)
            val l = DocxReader.read(dest).looks
            assertEquals(DocxReader.StyleLook("Times New Roman", 24, true, false, "center", 0), l.getValue("heading1"))
            assertEquals(DocxReader.StyleLook("Times New Roman", 24, true, false, "left", 0), l.getValue("heading2"))
            assertEquals(DocxReader.StyleLook("Times New Roman", 24, true, true, "left", 0), l.getValue("heading3"))
        }
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
