package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class ReviewTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-review").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }
    private val src = File(fixtures, "review-fixture.docx")

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    private fun save(ops: List<Op>, label: String, expected: List<String> = emptyList()): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"), expected)
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    private val lead = "D. Tracked changes: "

    @Test
    fun readsInsertionsAndDeletions() {
        val p = paragraphs(src)[4]
        assertEquals("${lead}this was inserted and .", p.text)
        val ins = p.revisions.single { it.kind == "ins" }
        assertEquals("10", ins.id)
        assertEquals("DOCX test", ins.author)
        assertEquals("this was inserted", p.runs.filter { it.rev == "10" }.joinToString("") { it.text })
        val del = p.revisions.single { it.kind == "del" }
        assertEquals("11", del.id)
        assertEquals("${lead}this was inserted and ".length, del.at)
        assertEquals("this was deleted", del.runs.joinToString("") { it.text })
    }

    @Test
    fun acceptsAndRejectsOneAtATime() {
        val rejectedDel = paragraphs(save(listOf(Op.Revision(4, "11", accept = false)), "reject-del"))[4]
        assertEquals("${lead}this was inserted and this was deleted.", rejectedDel.text)
        assertEquals(listOf("ins"), rejectedDel.revisions.map { it.kind })

        val rejectedIns = paragraphs(save(listOf(Op.Revision(4, "10", accept = false)), "reject-ins"))[4]
        assertEquals("$lead and .", rejectedIns.text)
        assertEquals(listOf("del"), rejectedIns.revisions.map { it.kind })
        assertEquals("$lead and ".length, rejectedIns.revisions.single().at)

        val acceptedIns = paragraphs(save(listOf(Op.Revision(4, "10", accept = true)), "accept-ins"))[4]
        assertEquals("${lead}this was inserted and .", acceptedIns.text)
        assertTrue(acceptedIns.runs.none { it.rev != null })
    }

    @Test
    fun acceptsAndRejectsAll() {
        val all = paragraphs(save(listOf(Op.Revision(-1, "*", accept = true)), "accept-all"))
        assertEquals("${lead}this was inserted and .", all[4].text)
        assertTrue(all.all { it.revisions.isEmpty() })
        val none = paragraphs(save(listOf(Op.Revision(-1, "*", accept = false)), "reject-all"))
        assertEquals("$lead and this was deleted.", none[4].text)
        assertTrue(none.all { it.revisions.isEmpty() })
    }

    @Test
    fun screenMustAgree() {
        val texts = paragraphs(src).map { it.text }.toMutableList()
        texts[4] = "${lead}this was inserted and this was deleted."
        save(listOf(Op.Revision(4, "11", accept = false)), "reject-del-expected", texts)
        texts[4] = "wrong"
        val failed = runCatching { save(listOf(Op.Revision(4, "11", accept = false)), "reject-del-wrong", texts) }
        assertTrue(failed.isFailure)
    }

    @Test(expected = IllegalStateException::class)
    fun unknownChangeIsRefused() {
        save(listOf(Op.Revision(4, "99", accept = true)), "unknown")
    }
}
