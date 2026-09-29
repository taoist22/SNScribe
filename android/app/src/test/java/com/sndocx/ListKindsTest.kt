package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Test
import java.io.File
import java.nio.file.Files

class ListKindsTest {
    @Test
    fun outlineListWithLevels() {
        val work = Files.createTempDirectory("sn-docx-lists").toFile()
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val dest = File(work, "outline.docx")
        DocxEditor.save(
            blank,
            listOf(
                Op.Text(0, 0, 0, "One"), Op.Split(0, 3), Op.Text(1, 0, 0, "Sub"), Op.Split(1, 3), Op.Text(2, 0, 0, "Two"),
                Op.ListItem(0, "outline", "o1"), Op.ListItem(1, "outline", "o1"), Op.ListItem(2, "outline", "o1"),
                Op.ListLevel(1, 1),
            ),
            dest,
            File(work, "tmp"),
        )
        File("build/test-out").mkdirs()
        dest.copyTo(File("build/test-out/lists-outline.docx"), overwrite = true)
        val ps = DocxReader.read(dest).blocks.filterIsInstance<DocxReader.Paragraph>()
        assertEquals(listOf("I.", "A.", "II."), ps.map { it.listLabel })
        assertEquals(listOf(0, 1, 0), ps.map { it.ilvl })
    }
}
