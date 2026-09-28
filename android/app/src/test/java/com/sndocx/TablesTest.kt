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

        val shrunk = save(grown, listOf(Op.TableRowDelete(0, 1)), "table-shrunk")
        val sh = tables(shrunk).single()
        assertEquals(2, sh.rows)
        assertEquals("new", sh.grid[1][1].pieces.single().text)
    }
}
