package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertNotNull
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipFile

/** Pictures: inserted as Word's inline pictures, read back with their file and size. */
class PicturesTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-pic").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }

    @Test
    fun insertsAFigureAndReadsItBack() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val png = File(fixtures, "ink-note.png")
        val dest = File(work, "figure.docx")
        // What figureOps makes for an APA figure after the first paragraph.
        val ops = listOf(
            Op.Text(0, 0, 0, "Intro."),
            Op.Split(0, 6),
            Op.Text(1, 0, 0, "Figure 1"),
            Op.Split(1, 8),
            Op.Text(2, 0, 0, "A Title"),
            Op.Split(2, 7),
            Op.ImageAdd(3, 0, png.path, 914400, 457200, "A Title"),
        )
        DocxEditor.save(blank, ops, dest, File(work, "tmp"))
        dest.copyTo(File(keep, "figure.docx"), overwrite = true)

        val ps = DocxReader.read(dest).blocks.filterIsInstance<DocxReader.Paragraph>()
        assertEquals(listOf("Intro.", "Figure 1", "A Title", DocxReader.OBJECT.toString()), ps.map { it.text })
        val run = ps[3].runs.single()
        assertEquals("image", run.obj)
        assertEquals(914400L, run.cx)
        assertEquals(457200L, run.cy)
        val rel = assertNotNull(run.imageRel).let { run.imageRel!! }

        val pictures = DocxReader.extractImages(dest, File(work, "pics"))
        assertTrue(File(pictures.getValue(rel)).readBytes().contentEquals(png.readBytes()))
        ZipFile(dest).use { z ->
            val types = String(z.getInputStream(z.getEntry("[Content_Types].xml")).readBytes())
            assertTrue(types.contains("Extension=\"png\""))
            val xml = String(z.getInputStream(z.getEntry("word/document.xml")).readBytes())
            assertTrue(xml.contains("<wp:inline"))
            assertTrue(xml.contains("descr=\"A Title\""))
        }
    }
}
