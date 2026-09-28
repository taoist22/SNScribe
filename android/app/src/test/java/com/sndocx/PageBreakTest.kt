package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream

/** Word page breaks (w:br type="page") and "starts on a new page" (w:pageBreakBefore). */
class PageBreakTest {
    private val work: File = Files.createTempDirectory("sn-docx-pb").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }

    private val body = "<w:p><w:r><w:t>Page one text.</w:t></w:r></w:p>" +
        "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>" +
        "<w:p><w:r><w:t>Page two, line one.</w:t><w:br w:type=\"page\"/><w:t>Page three.</w:t></w:r></w:p>" +
        "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>References and more</w:t></w:r></w:p>"

    /** A blank document with [body] in place of its empty paragraph. */
    private fun doc(): File {
        val blank = File(work, "blank-${System.nanoTime()}.docx").also { DocxBlank.write(it) }
        val out = File(work, "pb-${System.nanoTime()}.docx")
        ZipOutputStream(out.outputStream()).use { zo ->
            ZipFile(blank).use { z ->
                for (e in z.entries()) {
                    var bytes = z.getInputStream(e).readBytes()
                    if (e.name == "word/document.xml") bytes = String(bytes).replace("<w:body><w:p/>", "<w:body>$body").toByteArray()
                    zo.putNextEntry(ZipEntry(e.name))
                    zo.write(bytes)
                    zo.closeEntry()
                }
            }
        }
        return out
    }

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    private fun save(src: File, ops: List<Op>, label: String): File {
        val dest = File(work, "$label-${System.nanoTime()}.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"))
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    private fun documentXml(f: File) = ZipFile(f).use { z -> String(z.getInputStream(z.getEntry("word/document.xml")).readBytes()) }

    @Test
    fun readsWordPageBreaksAsTheirOwnRuns() {
        val ps = paragraphs(doc())
        assertEquals("\n", ps[1].text)
        assertTrue(ps[1].runs.single().pageBreak)
        assertEquals("Page two, line one.\nPage three.", ps[2].text)
        assertEquals(listOf(false, true, false), ps[2].runs.map { it.pageBreak })
        assertTrue(ps[3].para.pageBreakBefore == true)
    }

    @Test
    fun deletingAPageBreakRemovesIt() {
        val src = doc()
        val out = save(src, listOf(Op.Text(1, 0, 1, ""), Op.Text(2, 19, 20, "")), "pagebreak-deleted")
        val ps = paragraphs(out)
        assertEquals("", ps[1].text)
        assertEquals("Page two, line one.Page three.", ps[2].text)
        assertFalse(documentXml(out).contains("w:type=\"page\""))
    }

    @Test
    fun onlyTheFirstHalfOfASplitStartsANewPage() {
        val out = save(doc(), listOf(Op.Split(3, 10)), "pagebreak-split")
        val ps = paragraphs(out)
        assertEquals("References", ps[3].text)
        assertTrue(ps[3].para.pageBreakBefore == true)
        assertEquals(" and more", ps[4].text)
        assertFalse(ps[4].para.pageBreakBefore == true)
    }

    @Test
    fun turningOffIsWrittenOut() {
        val out = save(doc(), listOf(Op.ParaProps(3, pageBreakBefore = false)), "pagebreak-off")
        assertFalse(paragraphs(out)[3].para.pageBreakBefore == true)
        assertTrue(documentXml(out).contains("<w:pageBreakBefore w:val=\"0\"/>"))
    }
}
