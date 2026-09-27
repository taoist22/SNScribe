package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertNotNull
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class FontsAndBlankTest {
    private val work: File = Files.createTempDirectory("sn-docx-fonts").toFile()
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    @Test
    fun readsFamilyAndStyleFromFontFiles() {
        val arial = File("/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf")
        if (arial.isFile) assertEquals(FontFiles.Info("Arial", bold = true, italic = true), FontFiles.info(arial))
        val collection = File("/System/Library/Fonts/Helvetica.ttc")
        if (collection.isFile) assertEquals("Helvetica", FontFiles.info(collection)?.family)
        assertEquals(null, FontFiles.info(File(fixtures, "probe-fixture.docx")))
    }

    @Test
    fun findsFontsUnderLikelyNames() {
        val src = File("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf")
        if (!src.isFile) return
        src.copyTo(File(work, "timesbd.ttf"))
        src.copyTo(File(work, "Georgia-Italic.otf"))
        assertEquals("timesbd.ttf", FontFiles.find("Times New Roman", 1, listOf(work.path))?.name)
        assertEquals("Georgia-Italic.otf", FontFiles.find("Georgia", 2, listOf(work.path))?.name)
        assertEquals(null, FontFiles.find("Times New Roman", 0, listOf(work.path)))
    }

    @Test
    fun aBlankDocumentCanBeWrittenInAndSaved() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val ps = paragraphs(blank)
        assertEquals(1, ps.size)
        assertEquals("", ps[0].text)
        val dest = File(work, "My notes.docx")
        DocxEditor.save(
            blank,
            listOf(
                Op.Text(0, 0, 0, "My notes"),
                Op.Style(0, "heading1"),
                Op.Split(0, 8),
                Op.Text(1, 0, 0, "First point"),
                Op.ListItem(1, "number", "n1"),
                Op.Split(1, 11),
                Op.Text(2, 0, 0, "Second point"),
                Op.RunStyle(2, 0, 6, "Georgia", 28),
            ),
            dest,
            File(work, "tmp"),
        )
        val after = paragraphs(dest)
        assertEquals(listOf("My notes", "First point", "Second point"), after.map { it.text })
        assertEquals("heading", after[0].kind)
        assertEquals(listOf("1.", "2."), after.drop(1).map { it.listLabel })
        assertEquals("Calibri", after[1].runs.first().font) // docDefaults
        val second = after[2].runs.first()
        assertEquals("Second", second.text)
        assertEquals("Georgia", second.font)
        assertEquals(28, second.size)
    }

    @Test
    fun namesNewDocumentsWithoutOverwriting() {
        File(work, "Draft.docx").writeText("x")
        assertEquals("Draft 2.docx", DocxBlank.freeName(work, "Draft").name)
        assertEquals("a-b.docx", DocxBlank.freeName(work, "a/b").name)
    }

    @Test
    fun setsPageSetupAndDefaultFont() {
        val blank = File(work, "page.docx").also { DocxBlank.write(it) }
        val dest = File(work, "page-out.docx")
        val saved = DocxEditor.save(
            blank,
            listOf(
                Op.PageSetup(width = 15840, height = 12240, landscape = true, top = 1800, right = 1800, bottom = 1800, left = 1800),
                Op.Defaults("Times New Roman", 24),
                Op.Text(0, 0, 0, "Hello"),
            ),
            dest,
            File(work, "tmp"),
        )
        assertTrue(saved.changedParts.containsAll(listOf("word/document.xml", "word/styles.xml")))
        val r = DocxReader.read(dest)
        assertEquals(DocxReader.PageSetup(15840, 12240, 1800, 1800, 1800, 1800, landscape = true), r.page)
        val run = r.blocks.filterIsInstance<DocxReader.Paragraph>()[0].runs[0]
        assertEquals("Times New Roman", run.font)
        assertEquals(24, run.size)
    }

    /** Which fonts and sizes CT's real documents resolve to (theme fonts included). */
    @Test
    fun reportsFontsOfRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        for (f in File(dir).listFiles { x -> x.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }) {
            val runs = paragraphs(f).flatMap { it.runs }
            val fonts = runs.groupingBy { it.font ?: "(none)" }.eachCount()
            val sizes = runs.groupingBy { it.size ?: 0 }.eachCount()
            println("${f.name}: fonts $fonts; sizes (half-points) $sizes")
            assertNotNull(runs)
        }
    }
}
