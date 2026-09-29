package com.sndocx

import com.sndocx.DocxEditor.Op
import com.sndocx.DocxReader.NotePiece
import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File
import java.nio.file.Files

class TablesTest {
    private val work: File = Files.createTempDirectory("sn-docx-tbl").toFile()
    private val keep = File("build/test-out").also { it.mkdirs() }

    private fun save(src: File, ops: List<Op>, label: String): File {
        val dest = File(work, "$label.docx")
        DocxEditor.save(src, ops, dest, File(work, "tmp"))
        dest.copyTo(File(keep, "$label.docx"), overwrite = true)
        return dest
    }

    private fun tables(f: File) = DocxReader.read(f).blocks.filterIsInstance<DocxReader.Table>()

    @Test
    fun insertFillAndReshapeATable() {
        val blank = File(work, "blank.docx").also { DocxBlank.write(it) }
        val made = save(
            blank,
            listOf(
                Op.Text(0, 0, 0, "After the table."),
                Op.TableInsert(0, 2, 3, header = true),
                Op.TableCell(0, 0, 0, listOf(NotePiece("Group", bold = true))),
                Op.TableCell(0, 1, 2, listOf(NotePiece("12.5"), NotePiece("\n"), NotePiece("note", italic = true))),
            ),
            "table-made",
        )
        val t = tables(made).single()
        assertEquals(0, t.index)
        assertEquals(2, t.rows)
        assertEquals(3, t.cols)
        assertEquals(listOf(NotePiece("Group", bold = true)), t.grid[0][0].pieces)
        assertEquals("12.5\nnote", t.grid[1][2].pieces.joinToString("") { it.text })
        assertEquals("After the table.", DocxReader.read(made).blocks.filterIsInstance<DocxReader.Paragraph>().single().text)

        val grown = save(made, listOf(Op.TableRowAdd(0, 1, below = true), Op.TableCell(0, 2, 1, listOf(NotePiece("new")))), "table-grown")
        val g = tables(grown).single()
        assertEquals(3, g.rows)
        assertEquals("new", g.grid[2][1].pieces.single().text)
        assertTrue(g.grid[2][2].pieces.isEmpty())

        val narrow = save(grown, listOf(Op.TableWidth(0, 50)), "table-narrow")
        val nw = tables(narrow).single()
        assertEquals(0.5, nw.widthFrac, 0.01)
        assertEquals("center", nw.align)
        assertTrue(kotlin.math.abs(nw.widths.sum() - 9360 / 2) <= 3)
        val inset = save(blank, listOf(Op.TableInsert(0, 2, 2, header = false, pct = 75)), "table-inset")
        assertEquals(0.75, tables(inset).single().widthFrac, 0.01)

        // A row added under the header has no header line of its own.
        val underHeader = save(made, listOf(Op.TableRowAdd(0, 0, below = true)), "table-under-header")
        val xml = java.util.zip.ZipFile(underHeader).use { z -> String(z.getInputStream(z.getEntry("word/document.xml")).readBytes()) }
        assertEquals(1, xml.split("<w:tr>").count { it.contains("tcBorders") })
        val noTable = save(made, listOf(Op.TableDelete(0)), "table-deleted")
        assertTrue(tables(noTable).isEmpty())

        val shrunk = save(grown, listOf(Op.TableRowDelete(0, 1)), "table-shrunk")
        val sh = tables(shrunk).single()
        assertEquals(2, sh.rows)
        assertEquals("new", sh.grid[1][1].pieces.single().text)
    }

    /** A cell with a link or picture is not rewritten (audit 2026-09-28: it lost them). */
    @Test
    fun richCellsAreNotEdited() {
        val blank = File(work, "blank-rich.docx").also { DocxBlank.write(it) }
        val made = save(blank, listOf(Op.TableInsert(0, 1, 2, header = false), Op.TableCell(0, 0, 1, listOf(NotePiece("plain")))), "table-rich-base")
        // Put a hyperlink into cell 0 by hand.
        val rich = File(work, "rich.docx")
        java.util.zip.ZipOutputStream(rich.outputStream()).use { zo ->
            java.util.zip.ZipFile(made).use { z ->
                for (e in z.entries()) {
                    var bytes = z.getInputStream(e).readBytes()
                    if (e.name == "word/document.xml") {
                        val xml = String(bytes)
                        val i = xml.indexOf("</w:pPr></w:p></w:tc>")
                        bytes = (xml.substring(0, i) + "</w:pPr><w:hyperlink w:anchor=\"x\"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p></w:tc>" + xml.substring(i + "</w:pPr></w:p></w:tc>".length)).toByteArray()
                    }
                    zo.putNextEntry(java.util.zip.ZipEntry(e.name))
                    zo.write(bytes)
                    zo.closeEntry()
                }
            }
        }
        val t = tables(rich).single()
        assertTrue(t.grid[0][0].rich)
        assertTrue(!t.grid[0][1].rich)
        val refused = runCatching { save(rich, listOf(Op.TableCell(0, 0, 0, listOf(NotePiece("x")))), "table-rich-refused") }
        assertTrue(refused.isFailure)
        // A plain cell next to it still edits, with a tab kept as a tab.
        val ok = save(rich, listOf(Op.TableCell(0, 0, 1, listOf(NotePiece("a\tb")))), "table-rich-ok")
        val xml = java.util.zip.ZipFile(ok).use { z -> String(z.getInputStream(z.getEntry("word/document.xml")).readBytes()) }
        assertTrue(xml.contains("<w:hyperlink"))
        assertTrue(xml.contains("<w:tab/>"))
    }
}
