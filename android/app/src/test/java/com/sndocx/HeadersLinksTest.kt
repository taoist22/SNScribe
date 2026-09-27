package com.sndocx

import com.sndocx.DocxEditor.Op
import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertNotNull
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipFile

class HeadersLinksTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")
    private val work: File = Files.createTempDirectory("sn-docx-hl").toFile()

    /** Saved results are also kept here, for schema validation outside the tests. */
    private val keep = File("build/test-out").also { it.mkdirs() }

    private fun paragraphs(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Paragraph>()

    private fun save(src: File, ops: List<Op>, label: String): Pair<File, DocxEditor.Saved> {
        val dest = File(work, "${src.nameWithoutExtension}-${System.nanoTime()}.docx")
        val saved = DocxEditor.save(src, ops, dest, File(work, "tmp"))
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest to saved
    }

    private fun entry(f: File, name: String): String? = ZipFile(f).use { z -> z.getEntry(name)?.let { String(z.getInputStream(it).readBytes()) } }

    private fun blank(): File = File(work, "blank-${System.nanoTime()}.docx").also { DocxBlank.write(it) }

    @Test
    fun addsHeaderAndFooterToABlankDocument() {
        val (dest, saved) = save(
            blank(),
            listOf(
                Op.HeaderFooter("header", "Reatherford", pageNumber = true, align = "right"),
                Op.HeaderFooter("footer", "", pageNumber = true, align = "center"),
            ),
            "blank-header-footer",
        )
        assertTrue(saved.changedParts.containsAll(listOf("word/header1.xml", "word/footer1.xml", "[Content_Types].xml", "word/_rels/document.xml.rels")))
        val r = DocxReader.read(dest)
        assertEquals(DocxReader.HeaderFooter("Reatherford", true, "right"), r.header)
        assertEquals(DocxReader.HeaderFooter("", true, "center"), r.footer)
        val doc = entry(dest, "word/document.xml")!!
        // Header reference before footer reference, both first in sectPr.
        assertTrue(Regex("<w:sectPr[^>]*><w:headerReference[^>]*/><w:footerReference").containsMatchIn(doc))
        assertTrue(entry(dest, "[Content_Types].xml")!!.contains("/word/header1.xml"))
    }

    @Test
    fun rewritingAHeaderKeepsOnePart() {
        val (first, _) = save(blank(), listOf(Op.HeaderFooter("header", "One", true, "right")), "header-once")
        val (second, saved) = save(first, listOf(Op.HeaderFooter("header", "Two", false, "left")), "header-twice")
        assertEquals(listOf("word/header1.xml"), saved.changedParts.filter { it.contains("header") })
        assertEquals(DocxReader.HeaderFooter("Two", false, "left"), DocxReader.read(second).header)
        ZipFile(second).use { z -> assertFalse(z.entries().asSequence().any { it.name == "word/header2.xml" }) }
    }

    @Test
    fun linksAndUnlinks() {
        val src = File(fixtures, "probe-fixture.docx")
        val p1 = paragraphs(src)[1].text
        val start = p1.indexOf("Plain")
        val (linked, _) = save(src, listOf(Op.Link(1, start, start + 5, "https://doi.org/10.1037/0000165-000")), "linked")
        val runs = paragraphs(linked)[1].runs
        assertEquals("Plain", runs.filter { it.link }.joinToString("") { it.text })
        assertEquals(p1, paragraphs(linked)[1].text)
        val rels = entry(linked, "word/_rels/document.xml.rels")!!
        assertTrue(rels.contains("Target=\"https://doi.org/10.1037/0000165-000\" TargetMode=\"External\""))

        // A caret at the link's end removes it, as the screen does.
        val (unlinked, _) = save(linked, listOf(Op.Unlink(1, start + 5, start + 5)), "unlinked")
        assertFalse(paragraphs(unlinked)[1].runs.any { it.link && it.text.contains("Plain") })
        assertEquals(p1, paragraphs(unlinked)[1].text)
    }

    @Test(expected = IllegalStateException::class)
    fun refusesToLinkInsideALink() {
        val src = File(fixtures, "probe-fixture.docx")
        val (linked, _) = save(src, listOf(Op.Link(1, 0, 5, "https://example.com")), "linked-twice")
        save(linked, listOf(Op.Link(1, 2, 8, "https://example.org")), "linked-twice-2")
    }

    @Test
    fun templateBecomesADocument() {
        val dotx = File(work, "t.dotx")
        // A blank document declared a template, as Word writes a .dotx.
        val b = blank()
        java.util.zip.ZipOutputStream(dotx.outputStream()).use { out ->
            ZipFile(b).use { z ->
                for (e in z.entries()) {
                    var bytes = z.getInputStream(e).readBytes()
                    if (e.name == "[Content_Types].xml") {
                        bytes = String(bytes).replace("document.main+xml", "template.main+xml").toByteArray()
                    }
                    out.putNextEntry(java.util.zip.ZipEntry(e.name))
                    out.write(bytes)
                    out.closeEntry()
                }
            }
        }
        assertTrue(entry(dotx, "[Content_Types].xml")!!.contains("template.main+xml"))
        val doc = File(work, "from-template.docx")
        DocxBlank.fromTemplate(dotx, doc)
        val types = entry(doc, "[Content_Types].xml")!!
        assertTrue(types.contains("document.main+xml"))
        assertFalse(types.contains("template.main+xml"))
        assertNotNull(DocxReader.read(doc))
        doc.copyTo(File(keep, "from-template.docx"), overwrite = true)
    }

    /** On CT's real documents (DOCX_SAMPLES): read the header, then set header and footer. */
    @Test
    fun headersOnRealDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedBy { it.name }
        for ((i, f) in files.withIndex()) {
            val before = DocxReader.read(f)
            println("${f.name}: header=${before.header} footer=${before.footer}")
            val (dest, saved) = save(
                f,
                listOf(
                    Op.HeaderFooter("header", "Reatherford", true, "right"),
                    Op.HeaderFooter("footer", "", false, "center"),
                ),
                "real-$i",
            )
            println("  → ${saved.changedParts} ${saved.notes}")
            val after = DocxReader.read(dest)
            assertEquals("Reatherford", after.header!!.text)
            assertTrue(after.header!!.pageNumber)
            assertEquals("right", after.header!!.align)
            // Logos and tables stay: the part keeps every drawing and table it had.
            val names = ZipFile(f).use { z -> z.entries().asSequence().map { it.name }.filter { Regex("word/(header|footer)\\d*\\.xml").matches(it) }.toList() }
            for (n in names) {
                for (tag in listOf("<w:drawing>", "<w:tbl>")) {
                    assertEquals("$n $tag", Regex(Regex.escape(tag)).findAll(entry(f, n)!!).count(), Regex(Regex.escape(tag)).findAll(entry(dest, n)!!).count())
                }
            }
            assertEquals(before.blocks.size, after.blocks.size)
        }
    }
}
