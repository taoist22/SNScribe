package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

/** The edit sequence domain/citations.ts referenceOps makes, as the writer receives it. */
class CitationsTest {
    private val work: File = Files.createTempDirectory("sn-docx-cite").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }

    @Test
    fun citationAndNewReferenceList() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val body = "Mindfulness helps."
        val cited = " (Smith & Lee, 2020)"
        val entry = "Smith, J., & Lee, K. (2020). Mindfulness and test anxiety. Journal of Educational Psychology, 112(3), 455–467."
        val journal = entry.indexOf("Journal")
        val volume = entry.indexOf("112")
        val bodyAfter = body.substring(0, body.length - 1) + cited + "."
        val ops = listOf(
            Op.Text(0, 0, 0, body),
            Op.Text(0, body.length - 1, body.length - 1, cited),
            Op.Split(0, bodyAfter.length),
            Op.Text(1, 0, 0, "References"),
            Op.Style(1, "normal"),
            Op.ParaProps(1, "center", 480, "auto", 0, 0, 0, true),
            Op.Format(1, 0, 10, "b", true),
            Op.Format(1, 0, 10, "i", false),
            Op.Split(1, 10),
            Op.Text(2, 0, 0, entry),
            Op.Style(2, "normal"),
            Op.ParaProps(2, "left", 480, "auto", 0, 0, -720, false),
            Op.Format(2, 0, entry.length, "b", false),
            Op.Format(2, 0, entry.length, "i", false),
            Op.Format(2, journal, journal + "Journal of Educational Psychology".length, "i", true),
            Op.Format(2, volume, volume + 3, "i", true),
        )
        val dest = File(work, "cited.docx")
        DocxEditor.save(blank, ops, dest, File(work, "tmp"), listOf(bodyAfter, "References", entry))
        dest.copyTo(File(keep, "cite-new-list.docx"), overwrite = true)
        val ps = DocxReader.read(dest).blocks.filterIsInstance<DocxReader.Paragraph>()
        assertEquals(listOf(bodyAfter, "References", entry), ps.map { it.text })
        assertTrue(ps[1].para.pageBreakBefore == true)
        assertEquals("center", ps[1].align)
        assertTrue(ps[1].runs.all { it.bold })
        assertFalse(ps[2].para.pageBreakBefore == true)
        assertEquals(-720, ps[2].para.first)
        assertEquals("Journal of Educational Psychology112", ps[2].runs.filter { it.italic }.joinToString("") { it.text })
        assertFalse(ps[2].runs.any { it.bold })
    }
}
