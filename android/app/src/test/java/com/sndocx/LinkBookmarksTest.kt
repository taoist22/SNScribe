package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class LinkBookmarksTest {
    @Test
    fun hiddenBookmarkAroundWordsReadsBackAsMarks() {
        val work = Files.createTempDirectory("sn-docx-bm").toFile()
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val dest = File(work, "bm.docx")
        DocxEditor.save(
            blank,
            listOf(
                Op.Text(0, 0, 0, "Some words here. Next line"),
                Op.Split(0, 16),
                Op.BookmarkAdd(0, 5, 1, 5, "_sns_q_abc123"),
            ),
            dest,
            File(work, "tmp"),
        )
        File("build/test-out").mkdirs()
        dest.copyTo(File("build/test-out/bookmark-link.docx"), overwrite = true)
        val ps = DocxReader.read(dest).blocks.filterIsInstance<DocxReader.Paragraph>()
        assertEquals(listOf("Some words here.", " Next line"), ps.map { it.text })
        assertEquals(listOf(DocxReader.Mark("_sns_q_abc123", "start", 5)), ps[0].marks)
        assertEquals(listOf(DocxReader.Mark("_sns_q_abc123", "end", 5)), ps[1].marks)
        val bad = runCatching { DocxEditor.save(dest, listOf(Op.BookmarkAdd(0, 0, 0, 1, "_sns_q_abc123")), File(work, "b2.docx"), File(work, "tmp2")) }
        assertTrue(bad.isFailure)
    }
}
