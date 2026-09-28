package com.sndocx

import com.sndocx.DocxEditor.Op
import com.sndocx.DocxReader.NotePiece
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipFile

class FootnotesTest {
    private val work: File = Files.createTempDirectory("sn-docx-fn").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }

    private fun save(src: File, ops: List<Op>, label: String): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"))
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    private fun entry(f: File, name: String) = ZipFile(f).use { z -> z.getEntry(name)?.let { String(z.getInputStream(it).readBytes()) } }

    @Test
    fun addEditAndDeleteAFootnote() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val chicago = listOf(NotePiece("Jane Smith, "), NotePiece("A Book of Notes", italic = true), NotePiece(" (Chicago: Press, 2020), 23."))
        val added = save(blank, listOf(Op.Text(0, 0, 0, "A claim."), Op.FootnoteAdd(0, 8, 1, chicago)), "footnote-added")

        val r = DocxReader.read(added)
        val p = r.blocks.filterIsInstance<DocxReader.Paragraph>().first()
        assertEquals("A claim.${DocxReader.OBJECT}", p.text)
        assertEquals("1", p.runs.last().noteId)
        assertEquals(listOf(DocxReader.Footnote("1", chicago)), r.footnotes)
        assertTrue(entry(added, "[Content_Types].xml")!!.contains("/word/footnotes.xml"))
        assertTrue(entry(added, "word/styles.xml")!!.contains("footnote reference"))

        val edited = save(added, listOf(Op.FootnoteSet(1, listOf(NotePiece("Smith, "), NotePiece("Notes", italic = true), NotePiece(", 24.")))), "footnote-edited")
        assertEquals("Smith, Notes, 24.", DocxReader.read(edited).footnotes.single().text)

        val deleted = save(edited, listOf(Op.FootnoteDelete(1)), "footnote-deleted")
        val after = DocxReader.read(deleted)
        assertEquals("A claim.", after.blocks.filterIsInstance<DocxReader.Paragraph>().first().text)
        assertTrue(after.footnotes.isEmpty())
        assertFalse(entry(deleted, "word/document.xml")!!.contains("footnoteReference"))
    }
}
