package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipFile

class CommentsTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-comments").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }
    private val review = File(fixtures, "review-fixture.docx")

    private fun save(src: File, ops: List<Op>, label: String): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"), DocxReader.read(src).blocks.filterIsInstance<DocxReader.Paragraph>().map { it.text })
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    private fun entry(f: File, name: String) = ZipFile(f).use { z -> z.getEntry(name)?.let { String(z.getInputStream(it).readBytes()) } }

    @Test
    fun readsCommentsThreadsAndAnchors() {
        val r = DocxReader.read(review)
        assertEquals(listOf("0", "1", "2"), r.comments.map { it.id })
        assertEquals("A typed comment.", r.comments[0].text)
        assertEquals("0", r.comments[1].parent)
        assertEquals(1, r.comments[2].pictures)
        val p1 = r.blocks.filterIsInstance<DocxReader.Paragraph>()[1]
        val start = p1.marks.first { it.id == "0" && it.kind == "start" }
        val end = p1.marks.first { it.id == "0" && it.kind == "end" }
        assertEquals(3, start.at)
        assertEquals(p1.text.length, end.at)
        assertTrue(p1.marks.any { it.id == "1" && it.kind == "ref" })
    }

    @Test
    fun addsRepliesAndDeletes() {
        val ops = listOf(
            Op.CommentAdd(3, 4, 3, 4, 10, "Check this\nsecond line", "CT", "CT", "2026-09-26T21:00:00Z"),
            Op.CommentAdd(4, 0, 0, 0, 0, "Agreed.", "CT", "CT", "2026-09-26T21:01:00Z", parent = 0),
            Op.CommentDelete(listOf(2)),
        )
        val out = save(review, ops, "comments-edit")
        val r = DocxReader.read(out)
        assertEquals(listOf("0", "1", "3", "4"), r.comments.map { it.id })
        assertEquals("Check this\nsecond line", r.comments.first { it.id == "3" }.text)
        assertEquals("0", r.comments.first { it.id == "4" }.parent)
        val ps = r.blocks.filterIsInstance<DocxReader.Paragraph>()
        assertEquals(3, ps[4].marks.first { it.id == "3" && it.kind == "start" }.at)
        assertEquals(10, ps[4].marks.first { it.id == "3" && it.kind == "end" }.at)
        assertTrue(ps[1].marks.any { it.id == "4" && it.kind == "start" })
        assertTrue(ps.none { p -> p.marks.any { it.id == "2" } })
        // The deleted picture comment's image relationship stays harmless; the text is unchanged.
        assertEquals(DocxReader.read(review).blocks.filterIsInstance<DocxReader.Paragraph>().map { it.text }, ps.map { it.text })
    }

    @Test
    fun firstCommentCreatesThePart() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val out = save(blank, listOf(Op.CommentAdd(0, 0, 0, 0, 0, "Start here", "CT", "CT", "2026-09-26T21:00:00Z")), "comment-blank")
        assertTrue(entry(out, "[Content_Types].xml")!!.contains("/word/comments.xml"))
        assertEquals("Start here", DocxReader.read(out).comments.single().text)
        // A reply creates commentsExtended.
        val out2 = save(out, listOf(Op.CommentAdd(1, 0, 0, 0, 0, "Reply", "CT", "CT", "2026-09-26T21:00:00Z", parent = 0)), "comment-blank-reply")
        assertEquals("0", DocxReader.read(out2).comments.first { it.id == "1" }.parent)
    }
}
