package com.sndocx

import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File

class DocxReaderTest {
    private val fixtures = File(System.getProperty("fixtures.dir") ?: "fixtures")

    private fun paragraphs(r: DocxReader.Result) = r.blocks.filterIsInstance<DocxReader.Paragraph>()

    @Test
    fun readsTheProbeFixture() {
        val r = DocxReader.read(File(fixtures, "probe-fixture.docx"))
        val ps = paragraphs(r)

        assertEquals("heading", ps[0].kind)
        assertEquals(1, ps[0].level)
        assertEquals("Probe Fixture Heading", ps[0].text)

        val mixed = ps[1].runs
        assertTrue(mixed.any { it.text == "a bold phrase" && it.bold && !it.italic })
        assertTrue(mixed.any { it.text == "an italic phrase" && it.italic && !it.bold })

        // The footnote reference is one U+FFFC, superscript, inside the paragraph text.
        val note = ps[2].runs.single { it.obj == "note" }
        assertEquals(DocxReader.OBJECT.toString(), note.text)
        assertTrue(note.superscript)
        assertTrue(ps[2].text.contains("make it bold${DocxReader.OBJECT}, then"))

        assertTrue(ps[3].runs.any { it.text == "bold italic" && it.bold && it.italic })

        // Table between paragraphs 3 and 4; paragraph indexes skip it (they count w:p only).
        val tableAt = r.blocks.indexOfFirst { it is DocxReader.Table }
        val table = r.blocks[tableAt] as DocxReader.Table
        assertEquals(2, table.rows)
        assertEquals(2, table.cols)
        assertEquals(3, (r.blocks[tableAt - 1] as DocxReader.Paragraph).index)
        assertEquals(4, (r.blocks[tableAt + 1] as DocxReader.Paragraph).index)

        assertEquals("image", ps[4].runs.single().obj)
        assertEquals(1, r.report.images)
        assertEquals(1, r.report.tables)
        assertEquals(0, r.report.trackedChanges)
    }

    /**
     * Not an assertion test: dumps every .docx in DOCX_SAMPLES in a readable form, so a
     * person can compare it with the document in Word. Skipped when DOCX_SAMPLES is unset.
     */
    @Test
    fun dumpsSampleDocuments() {
        val dir = System.getProperty("docx.samples").orEmpty()
        if (dir.isEmpty()) return
        val files = File(dir).listFiles { f -> f.name.endsWith(".docx") && !f.name.startsWith("~") }.orEmpty().sortedBy { it.name }
        for (f in files) {
            val t0 = System.nanoTime()
            val r = runCatching { DocxReader.read(f) }
            val ms = (System.nanoTime() - t0) / 1_000_000
            println("\n===== ${f.name} (${ms} ms)")
            r.onFailure { println("FAILED: $it") }
            r.onSuccess { res ->
                println("report: ${res.report}")
                for (b in res.blocks) println(describe(b))
            }
        }
    }

    private fun describe(b: DocxReader.Block): String = when (b) {
        is DocxReader.Paragraph -> {
            val head = buildString {
                append("p${b.index}")
                if (b.kind != "body") append(" [${b.kind}${if (b.level > 0) " ${b.level}" else ""}]")
                if (b.align != "left") append(" {${b.align}}")
                if (b.indentTwips > 0) append(" >${b.indentTwips}")
                b.listLabel?.let { append(" «$it»") }
            }
            val body = b.runs.joinToString("") { r ->
                val flags = listOfNotNull(
                    "B".takeIf { r.bold }, "I".takeIf { r.italic }, "U".takeIf { r.underline && !r.link },
                    "S".takeIf { r.strike }, "H".takeIf { r.highlight }, "L".takeIf { r.link }, "^".takeIf { r.superscript },
                ).joinToString("")
                val text = if (r.obj != null) "[${r.obj}]" else r.text.replace("\n", "⏎").replace("\t", "⇥")
                if (flags.isEmpty()) text else "<$flags:$text>"
            }
            "$head: ${body.take(220)}${if (body.length > 220) "…" else ""}"
        }
        is DocxReader.Table -> "TABLE ${b.rows}×${b.cols}: ${b.preview.take(80)}"
        is DocxReader.Protected -> "PROTECTED ${b.what}: ${b.preview.take(80)}"
    }
}
