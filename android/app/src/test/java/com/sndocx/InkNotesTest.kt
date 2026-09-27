package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipFile

class InkNotesTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-ink").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }
    private val png = File(fixtures, "ink-note.png")

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()
    private fun names(f: File) = ZipFile(f).use { z -> z.entries().asSequence().map { it.name }.toSet() }

    private fun save(src: File, ops: List<Op>, label: String, expected: List<String>): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"), expected)
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    @Test
    fun addsReadsAndDeletesAMarginNote() {
        val src = File(fixtures, "probe-fixture.docx")
        val before = paragraphs(src).map { it.text }
        val at = before[1].indexOf("bold")
        val want = before.toMutableList().also { it[1] = it[1].substring(0, at) + DocxReader.OBJECT + it[1].substring(at) }
        val withNote = save(src, listOf(Op.InkAdd(1, at, "n1", png.path, 600, 260)), "ink-note", want)

        val run = paragraphs(withNote)[1].runs.single { it.obj == "ink" }
        assertEquals("n1", run.ink)
        val part = names(withNote).single { it.startsWith("word/media/sndocx-ink-") }
        assertTrue(ZipFile(withNote).use { z -> String(z.getInputStream(z.getEntry("[Content_Types].xml")).readBytes()) }.contains("Extension=\"png\""))
        val out = DocxReader.extractInk(withNote, File(work, "ink"))
        assertEquals(png.readBytes().toList(), File(out.getValue("n1")).readBytes().toList())

        val gone = save(withNote, listOf(Op.InkDelete(1, "n1")), "ink-note-deleted", before)
        assertTrue(paragraphs(gone)[1].runs.none { it.obj == "ink" })
        assertFalse(part in names(gone))
        assertFalse(ZipFile(gone).use { z -> String(z.getInputStream(z.getEntry("word/_rels/document.xml.rels")).readBytes()) }.contains(part.removePrefix("word/")))
    }

    @Test
    fun notesBesideOneParagraphStackAndNeverOverlap() {
        val src = File(fixtures, "probe-fixture.docx")
        val before = paragraphs(src).map { it.text }
        val want = before.toMutableList()
        fun anchorAt(i: Int, at: Int) { want[i] = want[i].substring(0, at) + DocxReader.OBJECT + want[i].substring(at) }
        anchorAt(1, 0); anchorAt(1, 3); anchorAt(2, 0)
        val out = save(
            src,
            listOf(Op.InkAdd(1, 0, "a", png.path, 600, 260), Op.InkAdd(1, 3, "b", png.path, 600, 260), Op.InkAdd(2, 0, "c", png.path, 600, 260)),
            "ink-three",
            want,
        )
        val xml = ZipFile(out).use { z -> String(z.getInputStream(z.getEntry("word/document.xml")).readBytes()) }
        val offsets = Regex("""<wp:positionV relativeFrom="line"><wp:posOffset>(\d+)</wp:posOffset>""").findAll(xml).map { it.groupValues[1].toLong() }.toList()
        assertEquals(3, offsets.size)
        // In document order: note a, note b (stacked below a, same paragraph), note c (its own paragraph).
        assertEquals(0L, offsets[0])
        assertTrue(offsets[1] > 0)
        assertEquals(0L, offsets[2])
        assertFalse(xml.contains("allowOverlap=\"1\""))
    }

    @Test
    fun addAndDeleteInOneSessionLeavesNoPicture() {
        val src = File(fixtures, "probe-fixture.docx")
        val before = paragraphs(src).map { it.text }
        val out = save(src, listOf(Op.InkAdd(0, 0, "n2", png.path, 600, 260), Op.InkDelete(0, "n2")), "ink-add-delete", before)
        assertTrue(names(out).none { it.startsWith("word/media/sndocx-ink-") })
    }
}
