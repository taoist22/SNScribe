package com.sndocx

import org.w3c.dom.Document
import org.w3c.dom.Element
import org.w3c.dom.Node
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.File
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import javax.xml.parsers.DocumentBuilderFactory

/**
 * Reads a .docx into a display model. Pure Kotlin + javax.xml, no React Native, so it runs
 * in JVM unit tests against real documents.
 *
 * The model is for reading and, later, for addressing edits:
 *  - Paragraph.index is the paragraph's ordinal among the body's top-level w:p elements —
 *    the address the writer (sn-docx-probe's DocxFileModule) edits by.
 *  - Paragraph text is the concatenation of its runs' text, built by [textOf]: every inline
 *    object (image, note reference, embedded object) is one U+FFFC so character offsets stay
 *    stable between the display and the XML. The writer must count the same way.
 *
 * Nothing here edits. Anything the reader cannot show faithfully becomes a labelled block or
 * a U+FFFC run, never silently dropped text.
 */
object DocxReader {
    const val W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    const val OBJECT = '￼'
    const val DOCUMENT_PART = "word/document.xml"
    const val DOCUMENT_RELS = "word/_rels/document.xml.rels"
    const val REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"

    /** The theme part (for theme fonts), through the document's relationships. */
    fun themePart(rels: Document?): String = relTarget(rels, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        ?: "word/theme/theme1.xml"

    fun relTarget(rels: Document?, type: String): String? {
        var n = rels?.documentElement?.firstChild
        while (n != null) {
            if (n is Element && n.localName == "Relationship" && n.getAttribute("Type") == type && n.getAttribute("TargetMode") != "External") {
                val target = n.getAttribute("Target")
                return if (target.startsWith("/")) target.removePrefix("/") else "word/" + target.removePrefix("./")
            }
            n = n.nextSibling
        }
        return null
    }

    const val R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    /** The part a document relationship id points at, or null (external or missing). */
    private fun relById(rels: Document?, id: String): String? {
        var n = rels?.documentElement?.firstChild
        while (n != null) {
            if (n is Element && n.localName == "Relationship" && n.getAttribute("Id") == id && n.getAttribute("TargetMode") != "External") {
                val target = n.getAttribute("Target")
                return if (target.startsWith("/")) target.removePrefix("/") else "word/" + target.removePrefix("./")
            }
            n = n.nextSibling
        }
        return null
    }

    /**
     * What DOCX can edit in a header or footer part. With only text in it, the whole part:
     * its text (without the page number itself), page number and alignment. With a logo or
     * table too ([HeaderFooter.other]), just its [editableLine].
     */
    fun headerFooter(part: Document): HeaderFooter {
        val other = hasOtherContent(part)
        val ps: List<Element> = if (other) {
            listOfNotNull(editableLine(part))
        } else {
            val all = part.getElementsByTagNameNS(W, "p")
            (0 until all.length).map { all.item(it) as Element }
        }
        val text = StringBuilder()
        var pageNumber = false
        var align = if (other) "right" else "left"
        for (p in ps) {
            val line = lineOf(p)
            if (line.page) pageNumber = true
            if (line.text.isNotBlank() || line.page) {
                child(p, "pPr")?.let { child(it, "jc") }?.getAttributeNS(W, "val")?.takeIf { it.isNotEmpty() }?.let { align = if (it == "both") "justify" else it }
            }
            if (line.text.isNotBlank()) {
                if (text.isNotEmpty()) text.append(" / ")
                text.append(line.text.trim())
            }
        }
        return HeaderFooter(text.toString(), pageNumber, align, other)
    }

    /**
     * The one line of a header or footer that holds more than text which DOCX may change:
     * the top-level paragraph with the page number, else an empty last paragraph, else
     * none (a new line goes at the bottom).
     */
    fun editableLine(part: Document): Element? {
        val top = elementChildren(part.documentElement).filter { it.localName == "p" }
        top.firstOrNull { lineOf(it).page }?.let { return it }
        val last = elementChildren(part.documentElement).lastOrNull()
        return last?.takeIf { it.localName == "p" && lineOf(it).text.isBlank() && !hasOther(it) }
    }

    private class Line(val text: String, val page: Boolean)

    /** A paragraph's shown text, leaving out a PAGE field's number, and whether it has one. */
    private fun lineOf(p: Element): Line {
        val t = StringBuilder()
        var page = false
        var inInstr = false
        var instr = StringBuilder()
        var skipping = false
        fun walk(n: org.w3c.dom.Node) {
            if (n !is Element) return
            when (n.localName) {
                "fldSimple" -> if (n.getAttributeNS(W, "instr").trim().startsWith("PAGE")) {
                    page = true
                    return
                }
                "fldChar" -> when (n.getAttributeNS(W, "fldCharType")) {
                    "begin" -> { inInstr = true; instr = StringBuilder() }
                    "separate" -> { inInstr = false; skipping = instr.trim().startsWith("PAGE"); if (skipping) page = true }
                    "end" -> { if (inInstr && instr.trim().startsWith("PAGE")) page = true; inInstr = false; skipping = false }
                }
                "instrText" -> if (inInstr) instr.append(n.textContent)
                "t" -> if (!inInstr && !skipping) t.append(n.textContent)
                "tab" -> if (!inInstr && !skipping && n.parentNode?.localName == "r") t.append(' ')
            }
            var c = n.firstChild
            while (c != null) { walk(c); c = c.nextSibling }
        }
        walk(p)
        return Line(t.toString(), page)
    }

    private val OTHER = listOf("drawing", "pict", "object", "tbl", "sdt", "txbxContent")

    private fun hasOther(e: Element) = OTHER.any { e.getElementsByTagNameNS(W, it).length > 0 } ||
        e.getElementsByTagNameNS("http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent").length > 0

    /** Whether a header or footer holds more than text and fields: pictures, shapes, tables, controls. */
    fun hasOtherContent(part: Document): Boolean = part.documentElement?.let(::hasOther) ?: false

    const val REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    const val REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"

    /** A piece of a footnote's text, with its own italics and bold (a Chicago note's book title). */
    data class NotePiece(val text: String, val italic: Boolean = false, val bold: Boolean = false)

    /** A footnote: its w:id and its text (paragraphs joined by "\n"), without the number. */
    data class Footnote(val id: String, val pieces: List<NotePiece>) {
        val text: String get() = pieces.joinToString("") { it.text }
    }

    /** The footnotes in a footnotes part, in order, leaving out Word's separators. */
    fun footnotes(part: Document): List<Footnote> =
        elementChildren(part.documentElement).filter { it.localName == "footnote" && it.namespaceURI == W && it.getAttributeNS(W, "type").let { t -> t.isEmpty() || t == "normal" } }.map { f ->
            val pieces = ArrayList<NotePiece>()
            for ((pi, p) in elementChildren(f).filter { it.localName == "p" }.withIndex()) {
                if (pi > 0) pieces.add(NotePiece("\n"))
                for (seg in segments(p)) {
                    if (!seg.isRun || elementChildren(seg.el).any { it.localName == "footnoteRef" }) continue
                    val text = runText(seg.el).replace(OBJECT.toString(), "")
                    if (text.isEmpty()) continue
                    val rPr = child(seg.el, "rPr")
                    val piece = NotePiece(text, isOn(rPr?.let { child(it, "i") }), isOn(rPr?.let { child(it, "b") }))
                    val last = pieces.lastOrNull()
                    if (last != null && last.italic == piece.italic && last.bold == piece.bold && last.text != "\n") {
                        pieces[pieces.size - 1] = last.copy(text = last.text + piece.text)
                    } else {
                        pieces.add(piece)
                    }
                }
            }
            // Word puts a space between the number and the text: not part of the text.
            if (pieces.isNotEmpty()) pieces[0] = pieces[0].copy(text = pieces[0].text.trimStart())
            Footnote(f.getAttributeNS(W, "id"), pieces.filter { it.text.isNotEmpty() })
        }
    const val REL_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
    const val W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
    const val W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

    /**
     * The comments, in file order. Replies and "done" come from commentsExtended, which
     * links a comment's last paragraph (w14:paraId) to its parent's.
     */
    fun comments(part: Document, extended: Document?): List<Comment> {
        val byPara = HashMap<String, String>() // last paragraph's paraId → comment id
        val raw = elementChildren(part.documentElement).filter { it.localName == "comment" && it.namespaceURI == W }
        for (c in raw) {
            lastParaId(c)?.let { byPara[it] = c.getAttributeNS(W, "id") }
        }
        val parentOf = HashMap<String, String>()
        val done = HashSet<String>()
        extended?.documentElement?.let { root ->
            for (e in elementChildren(root).filter { it.localName == "commentEx" }) {
                val id = byPara[e.getAttributeNS(W15, "paraId")] ?: continue
                byPara[e.getAttributeNS(W15, "paraIdParent")]?.let { parentOf[id] = it }
                if (e.getAttributeNS(W15, "done") == "1") done.add(id)
            }
        }
        return raw.map { c ->
            val id = c.getAttributeNS(W, "id")
            val paras = elementChildren(c).filter { it.localName == "p" }
            Comment(
                id = id,
                author = c.getAttributeNS(W, "author"),
                initials = c.getAttributeNS(W, "initials"),
                date = c.getAttributeNS(W, "date"),
                text = paras.joinToString("\n") { p -> segments(p).filter { it.isRun }.joinToString("") { runText(it.el) }.replace(OBJECT.toString(), "") }.trim(),
                parent = parentOf[id],
                done = id in done,
                pictures = c.getElementsByTagNameNS(W, "drawing").length + c.getElementsByTagNameNS(W, "pict").length,
            )
        }
    }

    /** The w14:paraId of a comment's last paragraph (what commentsExtended refers to). */
    fun lastParaId(comment: Element): String? =
        elementChildren(comment).lastOrNull { it.localName == "p" }?.getAttributeNS(W14, "paraId")?.takeIf { it.isNotEmpty() }

    const val WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    const val A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    /** docPr name of a handwritten note DOCX puts in the margin; the note id follows. */
    // Stays "DOCX ink note" after the rename to SNScribe: it marks notes already in documents.
    const val INK_NAME = "DOCX ink note "

    /** The note id of a w:drawing that is one of DOCX's handwritten margin notes, else null. */
    fun inkNoteId(drawing: Element): String? {
        val docPr = drawing.getElementsByTagNameNS(WP, "docPr")
        if (docPr.length == 0) return null
        val name = (docPr.item(0) as Element).getAttribute("name")
        return if (name.startsWith(INK_NAME)) name.removePrefix(INK_NAME).takeIf { it.isNotEmpty() } else null
    }

    /** The picture part of each handwritten margin note in [document]: note id → part name. */
    fun inkImages(document: Document, rels: Document?): Map<String, String> {
        val out = LinkedHashMap<String, String>()
        val drawings = document.getElementsByTagNameNS(W, "drawing")
        for (i in 0 until drawings.length) {
            val d = drawings.item(i) as Element
            val id = inkNoteId(d) ?: continue
            val blip = d.getElementsByTagNameNS(A, "blip").item(0) as? Element ?: continue
            relById(rels, blip.getAttributeNS(R_NS, "embed"))?.let { out[id] = it }
        }
        return out
    }

    data class EpubInfo(val title: String, val creators: List<String>, val date: String, val publisher: String)

    /**
     * An EPUB's own metadata (its OPF package file, found through META-INF/container.xml):
     * title, creators, date and publisher — enough to cite it as a book.
     */
    fun epubInfo(file: File): EpubInfo {
        ZipFile(file).use { zip ->
            val container = zip.getEntry("META-INF/container.xml")?.let { parse(readEntry(zip, it)) } ?: error("not an EPUB (no container.xml)")
            val rootfiles = container.getElementsByTagNameNS("*", "rootfile")
            val opfPath = (0 until rootfiles.length).map { rootfiles.item(it) as Element }
                .firstOrNull { it.getAttribute("media-type").contains("oebps-package") || it.getAttribute("full-path").endsWith(".opf") }
                ?.getAttribute("full-path") ?: error("no package file in container.xml")
            checkEntryName(opfPath)
            val opf = zip.getEntry(opfPath)?.let { parse(readEntry(zip, it)) } ?: error("missing $opfPath")
            val dc = "http://purl.org/dc/elements/1.1/"
            fun all(tag: String) = opf.getElementsByTagNameNS(dc, tag).let { l -> (0 until l.length).map { l.item(it).textContent.trim() }.filter { it.isNotEmpty() } }
            return EpubInfo(
                title = all("title").firstOrNull().orEmpty(),
                creators = all("creator"),
                date = all("date").firstOrNull().orEmpty(),
                publisher = all("publisher").firstOrNull().orEmpty(),
            )
        }
    }

    /** Picture formats Android can draw (Word's EMF/WMF and SVG can't be: those stay a placeholder). */
    private val RASTER = setOf("png", "jpg", "jpeg", "gif", "bmp", "webp")

    /**
     * Copies the document's pictures out of [file] into [dir], for the screen: relationship id →
     * file path. Only formats Android can draw; each at most 20 MB.
     */
    fun extractImages(file: File, dir: File): Map<String, String> {
        ZipFile(file).use { zip ->
            val rels = zip.getEntry(DOCUMENT_RELS)?.let { parse(readEntry(zip, it)) } ?: return emptyMap()
            val out = LinkedHashMap<String, String>()
            var n = rels.documentElement?.firstChild
            while (n != null) {
                if (n is Element && n.localName == "Relationship" && n.getAttribute("Type").endsWith("/image") && n.getAttribute("TargetMode") != "External") {
                    val target = n.getAttribute("Target")
                    val part = if (target.startsWith("/")) target.removePrefix("/") else "word/" + target.removePrefix("./")
                    val ext = part.substringAfterLast('.', "").lowercase()
                    val entry = zip.getEntry(part)
                    if (ext in RASTER && entry != null && entry.size in 1..20_000_000) {
                        checkEntryName(part)
                        dir.mkdirs()
                        val dest = File(dir, part.substringAfterLast('/').replace(Regex("[^A-Za-z0-9_.-]"), "_"))
                        if (!dest.exists() || dest.length() != entry.size) dest.writeBytes(readEntry(zip, entry))
                        out[n.getAttribute("Id")] = dest.path
                    }
                }
                n = n.nextSibling
            }
            return out
        }
    }

    /** Copies each handwritten note's picture out of [file] into [dir]; note id → file path. */
    fun extractInk(file: File, dir: File): Map<String, String> {
        ZipFile(file).use { zip ->
            val document = zip.getEntry(DOCUMENT_PART)?.let { parse(readEntry(zip, it)) } ?: return emptyMap()
            val rels = zip.getEntry(DOCUMENT_RELS)?.let { parse(readEntry(zip, it)) }
            val out = LinkedHashMap<String, String>()
            for ((id, part) in inkImages(document, rels)) {
                val entry = zip.getEntry(part) ?: continue
                checkEntryName(part)
                dir.mkdirs()
                val dest = File(dir, id.replace(Regex("[^A-Za-z0-9_-]"), "_") + ".png")
                dest.writeBytes(readEntry(zip, entry))
                out[id] = dest.path
            }
            return out
        }
    }

    fun alignOf(jc: String?): String = when (jc) {
        "center" -> "center"
        "right", "end" -> "right"
        "both", "distribute" -> "justify"
        else -> "left"
    }

    /** Where the document's list definitions live: its numbering relationship, else word/numbering.xml. */
    fun numberingPart(rels: Document?): String {
        val root = rels?.documentElement ?: return "word/numbering.xml"
        var n = root.firstChild
        while (n != null) {
            if (n is Element && n.localName == "Relationship" && n.getAttribute("Type") == REL_NUMBERING &&
                n.getAttribute("TargetMode") != "External"
            ) {
                val target = n.getAttribute("Target")
                return if (target.startsWith("/")) target.removePrefix("/") else "word/" + target.removePrefix("./")
            }
            n = n.nextSibling
        }
        return "word/numbering.xml"
    }
    const val MAX_DOCX_BYTES = 30L * 1024 * 1024
    private const val MAX_PART_BYTES = 40L * 1024 * 1024

    data class Run(
        val text: String,
        val bold: Boolean = false,
        val italic: Boolean = false,
        val underline: Boolean = false,
        val strike: Boolean = false,
        val highlight: Boolean = false,
        val link: Boolean = false,
        val superscript: Boolean = false,
        val subscript: Boolean = false,
        /** Word's highlight colour name ("yellow", "green", …) when [highlight]. */
        val highlightColor: String? = null,
        /** For a U+FFFC run: "image", "note" or "object". */
        val obj: String? = null,
        /** Inside a field or content control: shown and formattable, but its text is not editable. */
        val locked: Boolean = false,
        /** Font family name, theme fonts resolved; null = not set anywhere. */
        val font: String? = null,
        /** Size in half-points (w:sz); null = not set anywhere. */
        val size: Int? = null,
        /** Inside a tracked insertion (w:ins / w:moveTo): its w:id. */
        val rev: String? = null,
        /** A handwritten note DOCX put in the margin (obj "ink"): its note id. */
        val ink: String? = null,
        /** A Word page break (w:br w:type="page"): its text is "\n", and the next line starts a new page. */
        val pageBreak: Boolean = false,
        /** For obj "image": the picture's relationship id (a:blip r:embed) and its size in EMU (wp:extent). */
        val imageRel: String? = null,
        val cx: Long = 0,
        val cy: Long = 0,
        /** For obj "note": the footnote's w:id ("e" + id for an endnote). */
        val noteId: String? = null,
        /**
         * What the run sets itself (w:rPr and its character style), without what its
         * paragraph style gives it — the screen shows the style's part from the paragraph, so
         * a style change shows at once (CT: a new heading "did nothing" on screen because the
         * old style's size was baked into every run). The fields above stay fully resolved.
         */
        val ownBold: Boolean? = null,
        val ownItalic: Boolean? = null,
        val ownFont: String? = null,
        val ownSize: Int? = null,
    )

    /**
     * A tracked change in a paragraph's text. "ins": the runs with [Run.rev] == [id].
     * "del": text that is deleted but not yet accepted — not part of the paragraph's text;
     * it sits before character [at] and reads [runs]. [move]: a move (moveTo/moveFrom).
     */
    /** A comment anchor at text offset [at]: kind "start", "end" or "ref" (the comment's reference mark). */
    data class Mark(val id: String, val kind: String, val at: Int)

    /**
     * A comment (word/comments.xml). [parent]: the comment it replies to (commentsExtended);
     * [pictures]: how many pictures it holds (a handwritten note is one).
     */
    data class Comment(
        val id: String,
        val author: String,
        val initials: String,
        val date: String,
        val text: String,
        val parent: String? = null,
        val done: Boolean = false,
        val pictures: Int = 0,
    )

    data class Revision(
        val id: String,
        val kind: String,
        val author: String,
        val date: String,
        val at: Int = -1,
        val runs: List<Run> = emptyList(),
        val move: Boolean = false,
    )

    sealed class Block

    data class Paragraph(
        val index: Int,
        val styleId: String,
        /** "title", "subtitle", "heading" or "body". */
        val kind: String,
        /** Heading level 1–9; 0 otherwise. */
        val level: Int,
        /** "left", "center", "right" or "justify". */
        val align: String,
        val indentTwips: Int,
        val listLabel: String?,
        val runs: List<Run>,
        /** Its paragraph mark carries a section break (w:sectPr): it must not be joined away. */
        val sectionBreak: Boolean = false,
        /** The list it belongs to (w:numId, direct or from its style) and its level; null = none. */
        val numId: Int? = null,
        val ilvl: Int = 0,
        /** Spacing, first-line indent and page break, resolved through styles and defaults. */
        val para: ParaFmt = ParaFmt(),
        /** The font and size (half-points) text in this paragraph has unless it sets its own. */
        val baseFont: String? = null,
        val baseSize: Int? = null,
        /** Its style makes its text bold / italic (unless a run says otherwise). */
        val baseBold: Boolean = false,
        val baseItalic: Boolean = false,
        /** Tracked insertions and deletions in its text, in reading order. */
        val revisions: List<Revision> = emptyList(),
        /** In a quotation style (Quote / Intense Quote): body text, shown as a block quote. */
        val quote: Boolean = false,
        /** Its alignment / indent are set on the paragraph itself (a style change keeps them). */
        val ownAlign: Boolean = false,
        val ownIndent: Boolean = false,
        /** Where comments start, end and are referenced in its text. */
        val marks: List<Mark> = emptyList(),
    ) : Block() {
        val text: String get() = runs.joinToString("") { it.text }
    }

    data class Table(val rows: Int, val cols: Int, val preview: String) : Block()

    data class Protected(val what: String, val preview: String) : Block()

    data class Report(
        val paragraphs: Int,
        val tables: Int,
        val images: Int,
        val trackedChanges: Int,
        val comments: Int,
        val fields: Int,
        val contentControls: Int,
    )

    /** One list level: number format ("decimal", "bullet", …), label pattern ("%1."), first number. */
    data class ListLevel(val fmt: String, val text: String, val start: Int)

    /** A list (w:num): its levels, and start overrides that restart it. */
    data class ListDef(val levels: List<ListLevel?>, val starts: Map<Int, Int>)

    /** Page size and margins of the document's last section, in twips. */
    data class PageSetup(val width: Int, val height: Int, val top: Int, val right: Int, val bottom: Int, val left: Int, val landscape: Boolean)

    /** [lists] holds the definitions of every list a paragraph uses, so the screen can recount. */
    data class Result(
        val blocks: List<Block>,
        val report: Report,
        val lists: Map<Int, ListDef> = emptyMap(),
        val page: PageSetup? = null,
        val header: HeaderFooter? = null,
        val footer: HeaderFooter? = null,
        /** Tracked changes DOCX can't review yet: formatting, paragraph marks, sections, tables. */
        val otherRevisions: Int = 0,
        /** Style kind ("heading1", …, "normal") → how it looks here. */
        val looks: Map<String, StyleLook> = emptyMap(),
        val comments: List<Comment> = emptyList(),
        val footnotes: List<Footnote> = emptyList(),
    )

    /**
     * How a paragraph in one of the styles DOCX offers looks in this document: the style as
     * the file defines it, or as DocxEditor creates it when the file has none. The screen
     * applies it when a style is chosen, so it matches Word at once.
     */
    data class StyleLook(
        val font: String?,
        val size: Int?,
        val bold: Boolean,
        val italic: Boolean,
        val align: String,
        val indent: Int,
    )

    /** The default header or footer of the last section: its text, whether it shows a page number, alignment. */
    data class HeaderFooter(val text: String, val pageNumber: Boolean, val align: String, val other: Boolean = false)

    // ---------------------------------------------------------------- entry points

    fun read(file: File): Result {
        require(file.isFile) { "not a file: ${file.path}" }
        require(file.length() <= MAX_DOCX_BYTES) { "too large: ${file.length()} bytes" }
        ZipFile(file).use { zip ->
            fun part(name: String): Document? = zip.getEntry(name)?.let { parse(readEntry(zip, it)) }
            for (e in zip.entries()) checkEntryName(e.name)
            val document = part(DOCUMENT_PART) ?: throw IllegalStateException("no $DOCUMENT_PART in package")
            val rels = part(DOCUMENT_RELS)
            val result = read(document, part("word/styles.xml"), part(numberingPart(rels)), part(themePart(rels)))
            val all = document.getElementsByTagNameNS(W, "sectPr")
            val sects = (0 until all.length).map { all.item(it) as Element }.filter { (it.parentNode as? Element)?.localName != "sectPrChange" }
            // The last section's default header: its own, else the nearest section before it (Word's inheritance).
            fun hf(tag: String): HeaderFooter? {
                val ref = sects.asReversed().firstNotNullOfOrNull { s ->
                    elementChildren(s).firstOrNull { it.localName == tag && it.getAttributeNS(W, "type").let { t -> t.isEmpty() || t == "default" } }
                } ?: return null
                val target = relById(rels, ref.getAttributeNS(R_NS, "id")) ?: return null
                return headerFooter(part(target) ?: return null)
            }
            val comments = relTarget(rels, REL_COMMENTS)?.let { part(it) }?.let { c ->
                comments(c, relTarget(rels, REL_COMMENTS_EXTENDED)?.let { part(it) })
            }.orEmpty()
            val footnotes = relTarget(rels, REL_FOOTNOTES)?.let { part(it) }?.let { footnotes(it) }.orEmpty()
            return result.copy(header = hf("headerReference"), footer = hf("footerReference"), comments = comments, footnotes = footnotes)
        }
    }

    fun read(document: Document, styles: Document?, numbering: Document?, theme: Document? = null): Result {
        val ctx = Context(Styles(styles), Numbering(numbering), themeFonts(theme))
        val body = child(document.documentElement, "body") ?: return Result(emptyList(), ctx.report())
        val blocks = ArrayList<Block>()
        var paraIndex = 0
        for (el in elementChildren(body)) {
            if (el.namespaceURI != W) continue
            when (el.localName) {
                "p" -> blocks.add(ctx.paragraph(el, paraIndex++))
                "tbl" -> {
                    ctx.tables++
                    val rows = elementChildren(el).filter { it.localName == "tr" }
                    val cols = rows.maxOfOrNull { r -> elementChildren(r).count { it.localName == "tc" } } ?: 0
                    blocks.add(Table(rows.size, cols, preview(el)))
                    ctx.scan(el)
                }
                "sdt" -> {
                    ctx.contentControls++
                    blocks.add(Protected("Content control", preview(el)))
                    ctx.scan(el)
                }
                "customXml", "altChunk" -> {
                    blocks.add(Protected(if (el.localName == "altChunk") "Embedded document" else "Custom XML block", preview(el)))
                    ctx.scan(el)
                }
                // sectPr, bookmarkStart/End, proofErr and other markers carry no content.
            }
        }
        val used = blocks.filterIsInstance<Paragraph>().mapNotNull { it.numId }.toSet()
        return Result(
            blocks,
            ctx.report(),
            used.mapNotNull { id -> ctx.numbering.definition(id)?.let { id to it } }.toMap(),
            pageSetup(child(body, "sectPr")),
            otherRevisions = otherRevisions(document),
            looks = ctx.looks(),
        )
    }

    /** Formatting changes, and insertions or deletions of paragraph marks and table rows. */
    fun otherRevisions(document: Document): Int {
        var n = 0
        for (tag in listOf("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange")) {
            n += document.getElementsByTagNameNS(W, tag).length
        }
        for (tag in listOf("ins", "del")) {
            val all = document.getElementsByTagNameNS(W, tag)
            for (i in 0 until all.length) if ((all.item(i).parentNode as? Element)?.localName in setOf("rPr", "trPr")) n++
        }
        return n
    }

    // ---------------------------------------------------------------- paragraphs

    /** The theme's heading (major) and body (minor) Latin fonts. */
    private data class ThemeFonts(val major: String?, val minor: String?) {
        fun resolve(font: String?): String? = when (font) {
            "+major" -> major
            "+minor" -> minor
            else -> font
        }
    }

    private fun themeFonts(theme: Document?): ThemeFonts {
        val a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        fun latin(which: String): String? {
            val list = theme?.getElementsByTagNameNS(a, which) ?: return null
            if (list.length == 0) return null
            val latin = (list.item(0) as Element).getElementsByTagNameNS(a, "latin")
            return if (latin.length > 0) (latin.item(0) as Element).getAttribute("typeface").takeIf { it.isNotEmpty() } else null
        }
        return ThemeFonts(latin("majorFont"), latin("minorFont"))
    }

    private class Context(val styles: Styles, val numbering: Numbering, val theme: ThemeFonts) {
        var tables = 0
        var images = 0
        var tracked = 0
        var comments = 0
        var fields = 0
        var contentControls = 0
        var paragraphs = 0

        fun report() = Report(paragraphs, tables, images, tracked, comments, fields, contentControls)

        fun looks(): Map<String, StyleLook> {
            val normal = styles.paragraph("")
            fun look(info: StyleInfo) = StyleLook(
                font = theme.resolve(info.run.font),
                size = info.run.size,
                bold = info.run.bold == true,
                italic = info.run.italic == true,
                align = alignOf(info.jc),
                indent = info.indent,
            )
            val out = LinkedHashMap<String, StyleLook>()
            out["normal"] = look(normal)
            // What DocxEditor.StyleIds.idFor creates when the file lacks the style.
            val created = mapOf(
                "heading1" to ("heading 1" to look(normal).copy(bold = true, size = 32)),
                "heading2" to ("heading 2" to look(normal).copy(bold = true, size = 28)),
                "heading3" to ("heading 3" to look(normal).copy(bold = true, size = 24)),
                "title" to ("title" to look(normal).copy(bold = true, size = 56)),
                "quote" to ("quote" to look(normal).copy(italic = true, indent = 720)),
            )
            for ((kind, pair) in created) {
                val id = styles.idByName(pair.first)
                out[kind] = if (id != null) look(styles.paragraph(id)) else pair.second
            }
            return out
        }

        /** Counts features inside blocks the reader does not show (tables, controls). */
        fun scan(el: Element) {
            val all = el.getElementsByTagNameNS(W, "*")
            for (i in 0 until all.length) {
                when ((all.item(i) as Element).localName) {
                    "drawing", "pict" -> images++
                    "ins", "del", "moveFrom", "moveTo" -> tracked++
                    "commentRangeStart" -> comments++
                    "fldChar", "fldSimple" -> fields++
                }
            }
        }

        fun paragraph(p: Element, index: Int): Paragraph {
            paragraphs++
            val pPr = child(p, "pPr")
            val styleId = pPr?.let { child(it, "pStyle") }?.getAttributeNS(W, "val").orEmpty()
            val style = styles.paragraph(styleId)
            val outline = pPr?.let { child(it, "outlineLvl") }?.getAttributeNS(W, "val")?.toIntOrNull() ?: style.outline
            val name = style.name.lowercase()
            val kind = when {
                name == "title" -> "title"
                name == "subtitle" -> "subtitle"
                outline in 0..8 -> "heading"
                else -> "body"
            }
            val jc = pPr?.let { child(it, "jc") }?.getAttributeNS(W, "val") ?: style.jc
            val align = when (jc) {
                "center" -> "center"
                "right", "end" -> "right"
                "both", "distribute" -> "justify"
                else -> "left"
            }
            val numPr = pPr?.let { child(it, "numPr") }
            val numId = numPr?.let { child(it, "numId") }?.getAttributeNS(W, "val")?.toIntOrNull() ?: style.numId
            val ilvl = numPr?.let { child(it, "ilvl") }?.getAttributeNS(W, "val")?.toIntOrNull() ?: style.ilvl
            val label = if (numId != null && numId > 0) numbering.label(numId, ilvl) else null
            val indent = pPr?.let { child(it, "ind") }?.let { twips(it) }
                ?: (if (label != null) numbering.indent(numId!!, ilvl) else null)
                ?: style.indent

            val runs = ArrayList<Run>()
            collectRuns(p, style.run, runs)
            val revisions = ArrayList<Revision>()
            val marks = ArrayList<Mark>()
            collectRevisions(p, style.run, intArrayOf(0), revisions, marks)
            return Paragraph(
                index = index,
                styleId = styleId,
                kind = kind,
                level = if (kind == "heading") outline!! + 1 else 0,
                align = align,
                indentTwips = indent,
                listLabel = label,
                runs = mergeAdjacent(runs),
                sectionBreak = pPr?.let { child(it, "sectPr") } != null,
                numId = numId?.takeIf { it > 0 && label != null },
                ilvl = ilvl.coerceIn(0, 8),
                para = style.para.merge(ParaFmt.of(pPr)),
                baseFont = theme.resolve(style.run.font),
                baseSize = style.run.size,
                baseBold = style.run.bold == true,
                baseItalic = style.run.italic == true,
                revisions = revisions,
                marks = marks,
                quote = kind == "body" && (name == "quote" || name == "intense quote"),
                ownAlign = pPr?.let { child(it, "jc") } != null,
                ownIndent = pPr?.let { child(it, "ind") } != null,
            )
        }

        /**
         * Tracked changes in reading order, mirroring [walk]: insertions where they start,
         * deletions with their position in the paragraph's text ([at], kept in [offset]).
         */
        private fun collectRevisions(parent: Element, base: Fmt, offset: IntArray, out: MutableList<Revision>, marks: MutableList<Mark>) {
            fun meta(el: Element, kind: String, at: Int = -1, runs: List<Run> = emptyList()) = Revision(
                id = el.getAttributeNS(W, "id"),
                kind = kind,
                author = el.getAttributeNS(W, "author"),
                date = el.getAttributeNS(W, "date"),
                at = at,
                runs = runs,
                move = el.localName.startsWith("move"),
            )
            for (el in elementChildren(parent)) {
                if (el.namespaceURI != W) {
                    if (el.localName == "AlternateContent") offset[0]++
                    continue
                }
                when (el.localName) {
                    "r" -> {
                        offset[0] += runText(el).length
                        child(el, "commentReference")?.let { marks.add(Mark(it.getAttributeNS(W, "id"), "ref", offset[0])) }
                    }
                    "commentRangeStart" -> marks.add(Mark(el.getAttributeNS(W, "id"), "start", offset[0]))
                    "commentRangeEnd" -> marks.add(Mark(el.getAttributeNS(W, "id"), "end", offset[0]))
                    "ins", "moveTo" -> {
                        if (lengthOf(el) > 0) out.add(meta(el, "ins"))
                        collectRevisions(el, base, offset, out, marks)
                    }
                    "del", "moveFrom" -> {
                        val deleted = ArrayList<Run>()
                        for (i in 0 until el.getElementsByTagNameNS(W, "r").length) {
                            run(el.getElementsByTagNameNS(W, "r").item(i) as Element, base, link = false, locked = true, out = deleted, deleted = true)
                        }
                        if (deleted.isNotEmpty()) out.add(meta(el, "del", offset[0], mergeAdjacent(deleted)))
                    }
                    "hyperlink", "fldSimple", "smartTag", "customXml" -> collectRevisions(el, base, offset, out, marks)
                    "sdt" -> child(el, "sdtContent")?.let { collectRevisions(it, base, offset, out, marks) }
                }
            }
        }

        private fun collectRuns(p: Element, base: Fmt, out: MutableList<Run>) {
            val segs = segments(p) { marker ->
                when (marker.localName) {
                    "ins", "moveTo", "del", "moveFrom" -> tracked++
                    "fldSimple" -> fields++
                    "sdt" -> contentControls++
                    "commentRangeStart" -> comments++
                    "AlternateContent" -> images++
                }
            }
            for (seg in segs) {
                if (seg.isRun) run(seg.el, base, seg.link, seg.locked, out, rev = seg.rev) else out.add(Run(OBJECT.toString(), obj = "object", locked = true, rev = seg.rev))
            }
        }

        private fun run(r: Element, base: Fmt, link: Boolean, locked: Boolean, out: MutableList<Run>, rev: String? = null, deleted: Boolean = false) {
            val rPr = child(r, "rPr")
            val charStyle = rPr?.let { child(it, "rStyle") }?.getAttributeNS(W, "val")
            val own = styles.character(charStyle).merge(Fmt.of(rPr))
            val fmt = base.merge(own)
            val isLink = link || charStyle.equals("Hyperlink", ignoreCase = true)
            fun add(text: String, obj: String? = null, ink: String? = null, pageBreak: Boolean = false, picture: Element? = null, noteId: String? = null) {
                if (text.isEmpty()) return
                out.add(
                    Run(
                        text = text,
                        bold = fmt.bold == true,
                        italic = fmt.italic == true,
                        underline = fmt.underline == true || isLink,
                        strike = fmt.strike == true,
                        highlight = fmt.highlight != null && fmt.highlight != "none",
                        highlightColor = fmt.highlight?.takeIf { it != "none" },
                        link = isLink,
                        superscript = fmt.vertAlign == "superscript" || obj == "note",
                        subscript = fmt.vertAlign == "subscript" && obj != "note",
                        obj = obj,
                        locked = locked,
                        font = when (fmt.font) {
                            "+major" -> theme.major
                            "+minor" -> theme.minor
                            else -> fmt.font
                        },
                        size = fmt.size,
                        rev = rev,
                        ink = ink,
                        pageBreak = pageBreak,
                        imageRel = picture?.let { d -> (d.getElementsByTagNameNS(A, "blip").item(0) as? Element)?.getAttributeNS(R_NS, "embed")?.takeIf { it.isNotEmpty() } },
                        cx = picture?.let { d -> (d.getElementsByTagNameNS(WP, "extent").item(0) as? Element)?.getAttribute("cx")?.toLongOrNull() } ?: 0,
                        cy = picture?.let { d -> (d.getElementsByTagNameNS(WP, "extent").item(0) as? Element)?.getAttribute("cy")?.toLongOrNull() } ?: 0,
                        noteId = noteId,
                        ownBold = own.bold,
                        ownItalic = own.italic,
                        ownFont = theme.resolve(own.font),
                        ownSize = own.size,
                    ),
                )
            }
            val text = StringBuilder()
            for (c in elementChildren(r)) {
                val t = if (deleted && c.localName == "delText") c.textContent else textOf(c)
                if (t.isEmpty()) {
                    if (c.localName == "fldChar") fields++
                    continue
                }
                if (c.localName == "br" && c.getAttributeNS(W, "type") == "page") {
                    // A page break Word shows as one: its own run, so the screen can show it and page there.
                    add(text.toString())
                    text.clear()
                    add(t, pageBreak = true)
                } else if (t[0] == OBJECT) {
                    add(text.toString())
                    text.clear()
                    val ink = if (c.localName == "drawing") inkNoteId(c) else null
                    val kind = when {
                        ink != null -> "ink"
                        c.localName == "drawing" || c.localName == "pict" -> "image".also { images++ }
                        c.localName == "footnoteReference" || c.localName == "endnoteReference" -> "note"
                        else -> "object"
                    }
                    val noteId = when (c.localName) {
                        "footnoteReference" -> c.getAttributeNS(W, "id")
                        "endnoteReference" -> "e" + c.getAttributeNS(W, "id")
                        else -> null
                    }
                    add(t, kind, ink, picture = if (kind == "image" && c.localName == "drawing") c else null, noteId = noteId)
                } else {
                    text.append(t)
                }
            }
            add(text.toString())
        }
    }

    /**
     * One text-bearing piece of a paragraph, in reading order: a w:r run, or a foreign object
     * (mc:AlternateContent — a text box or shape) that counts as a single U+FFFC.
     */
    class Segment(val el: Element, val isRun: Boolean, val link: Boolean, var locked: Boolean = false, val rev: String? = null) {
        val length: Int get() = if (isRun) runText(el).length else 1
    }

    /**
     * THE definition of a paragraph's text, shared by the reader and the writer so their
     * character offsets always agree. Descends through containers Word shows inline
     * (hyperlinks, insertions, simple fields, smart tags, custom XML, inline content
     * controls) and skips deleted or moved-away text, which Word does not show either.
     * [onMarker] sees every container and marker passed on the way (for the report).
     */
    fun segments(p: Element, onMarker: (Element) -> Unit = {}): List<Segment> {
        val out = ArrayList<Segment>()
        walk(p, false, out, onMarker, locked = false)
        // Complex fields (fldChar begin … end) span several runs: all of them, the codes and
        // the shown result, are locked against text edits. Nesting is counted.
        var depth = 0
        for (seg in out) {
            if (!seg.isRun) continue
            var begins = 0
            var ends = 0
            for (c in elementChildren(seg.el)) {
                if (c.localName == "fldChar") when (c.getAttributeNS(W, "fldCharType")) {
                    "begin" -> begins++
                    "end" -> ends++
                }
            }
            if (depth > 0 || begins > 0 || ends > 0) seg.locked = true
            depth = maxOf(0, depth + begins - ends)
        }
        return out
    }

    /** [locked]: inside a simple field or content control — shown, formattable, not text-editable. */
    private fun walk(parent: Element, link: Boolean, out: MutableList<Segment>, onMarker: (Element) -> Unit, locked: Boolean, rev: String? = null) {
        for (el in elementChildren(parent)) {
            if (el.namespaceURI != W) {
                if (el.localName == "AlternateContent") {
                    onMarker(el)
                    out.add(Segment(el, false, link, locked, rev))
                }
                continue
            }
            when (el.localName) {
                "r" -> out.add(Segment(el, true, link, locked, rev))
                "hyperlink" -> walk(el, true, out, onMarker, locked, rev)
                "ins", "moveTo" -> {
                    onMarker(el)
                    walk(el, link, out, onMarker, locked, el.getAttributeNS(W, "id"))
                }
                "fldSimple" -> {
                    onMarker(el)
                    walk(el, link, out, onMarker, true, rev)
                }
                "smartTag", "customXml" -> walk(el, link, out, onMarker, locked, rev)
                "sdt" -> {
                    onMarker(el)
                    child(el, "sdtContent")?.let { walk(it, link, out, onMarker, true, rev) }
                }
                else -> onMarker(el) // del, moveFrom, commentRangeStart, bookmarks, proofErr, pPr …
            }
        }
    }

    /**
     * How many characters of paragraph text [el] (a paragraph child, at any depth) holds —
     * exactly what [segments] would count for it.
     */
    fun lengthOf(el: Element): Int = when {
        el.namespaceURI != W -> if (el.localName == "AlternateContent") 1 else 0
        el.localName == "r" -> runText(el).length
        el.localName in CONTAINERS -> elementChildren(el).sumOf { lengthOf(it) }
        el.localName == "sdt" -> child(el, "sdtContent")?.let { c -> elementChildren(c).sumOf { lengthOf(it) } } ?: 0
        else -> 0
    }

    /** Inline containers [walk] descends into (sdt is handled through its sdtContent). */
    val CONTAINERS = setOf("hyperlink", "ins", "moveTo", "fldSimple", "smartTag", "customXml")

    fun runText(r: Element): String = buildString { for (c in elementChildren(r)) append(textOf(c)) }

    /**
     * What one run child contributes to the paragraph's text. Breaks are '\n', tabs '\t',
     * inline objects one U+FFFC. Field codes (instrText) and deleted text contribute nothing.
     */
    fun textOf(c: Element): String = when (c.localName) {
        "t" -> c.textContent
        "tab", "ptab" -> "\t"
        "br", "cr" -> "\n"
        "noBreakHyphen" -> "‑"
        "softHyphen" -> ""
        "sym" -> OBJECT.toString()
        "drawing", "pict", "object", "footnoteReference", "endnoteReference" -> OBJECT.toString()
        else -> ""
    }

    /** Neighbouring runs with identical formatting become one (Word splits runs for revision ids). */
    private fun mergeAdjacent(runs: List<Run>): List<Run> {
        val out = ArrayList<Run>()
        for (r in runs) {
            val last = out.lastOrNull()
            if (last != null && last.obj == null && r.obj == null && last.copy(text = "") == r.copy(text = "")) {
                out[out.size - 1] = last.copy(text = last.text + r.text)
            } else {
                out.add(r)
            }
        }
        return out
    }

    // ---------------------------------------------------------------- formatting

    /** Run formatting at one level; null = not set here, inherit. */
    data class Fmt(
        val bold: Boolean? = null,
        val italic: Boolean? = null,
        val underline: Boolean? = null,
        val strike: Boolean? = null,
        /** Highlight colour name, "none" to switch an inherited one off; null = inherit. */
        val highlight: String? = null,
        /** "superscript", "subscript" or "baseline"; null = inherit. */
        val vertAlign: String? = null,
        /** A family name, or "+major" / "+minor" for the theme's heading / body font. */
        val font: String? = null,
        val size: Int? = null,
    ) {
        fun merge(over: Fmt) = Fmt(
            over.bold ?: bold,
            over.italic ?: italic,
            over.underline ?: underline,
            over.strike ?: strike,
            over.highlight ?: highlight,
            over.vertAlign ?: vertAlign,
            over.font ?: font,
            over.size ?: size,
        )

        companion object {
            fun of(rPr: Element?): Fmt {
                if (rPr == null) return Fmt()
                fun on(name: String): Boolean? = child(rPr, name)?.let { isOn(it) }
                val u = child(rPr, "u")?.getAttributeNS(W, "val")
                val hl = child(rPr, "highlight")?.getAttributeNS(W, "val")
                val va = child(rPr, "vertAlign")?.getAttributeNS(W, "val")
                return Fmt(
                    bold = on("b"),
                    italic = on("i"),
                    underline = u?.let { it.isNotEmpty() && it != "none" },
                    strike = on("strike") ?: on("dstrike"),
                    highlight = hl?.takeIf { it.isNotEmpty() },
                    vertAlign = va?.takeIf { it.isNotEmpty() },
                    font = child(rPr, "rFonts")?.let { fontOf(it) },
                    size = child(rPr, "sz")?.getAttributeNS(W, "val")?.toIntOrNull(),
                )
            }

            /** The Latin-text font of w:rFonts: a named font, else a theme font marker. */
            private fun fontOf(r: Element): String? {
                fun attr(name: String) = r.getAttributeNS(W, name).takeIf { it.isNotEmpty() }
                fun theme(v: String?) = v?.let { if (it.startsWith("major")) "+major" else "+minor" }
                return attr("ascii") ?: theme(attr("asciiTheme")) ?: attr("hAnsi") ?: theme(attr("hAnsiTheme"))
            }
        }
    }

    /**
     * Paragraph spacing and first-line indent at one level (null = inherit): space before
     * and after and line spacing in twips (line in 240ths of a line when [lineRule] is
     * "auto"), first-line indent in twips (negative = hanging), and a page break before.
     */
    data class ParaFmt(
        val before: Int? = null,
        val after: Int? = null,
        val line: Int? = null,
        val lineRule: String? = null,
        val first: Int? = null,
        val pageBreakBefore: Boolean? = null,
    ) {
        fun merge(over: ParaFmt) = ParaFmt(
            over.before ?: before,
            over.after ?: after,
            over.line ?: line,
            over.lineRule ?: lineRule,
            over.first ?: first,
            over.pageBreakBefore ?: pageBreakBefore,
        )

        companion object {
            fun of(pPr: Element?): ParaFmt {
                if (pPr == null) return ParaFmt()
                val sp = child(pPr, "spacing")
                val ind = child(pPr, "ind")
                fun attr(e: Element?, n: String) = e?.getAttributeNS(W, n)?.takeIf { it.isNotEmpty() }
                val hanging = attr(ind, "hanging")?.toIntOrNull()
                val firstLine = attr(ind, "firstLine")?.toIntOrNull()
                return ParaFmt(
                    before = attr(sp, "before")?.toIntOrNull(),
                    after = attr(sp, "after")?.toIntOrNull(),
                    line = attr(sp, "line")?.toIntOrNull(),
                    lineRule = attr(sp, "lineRule") ?: attr(sp, "line")?.let { "auto" },
                    first = hanging?.let { -it } ?: firstLine,
                    pageBreakBefore = child(pPr, "pageBreakBefore")?.let { isOn(it) },
                )
            }
        }
    }

    private class StyleInfo(
        val name: String = "",
        val outline: Int? = null,
        val jc: String? = null,
        val numId: Int? = null,
        val ilvl: Int = 0,
        val indent: Int = 0,
        val run: Fmt = Fmt(),
        val para: ParaFmt = ParaFmt(),
    )

    /** styles.xml, resolved through basedOn chains (cycle-safe). */
    private class Styles(doc: Document?) {
        private val raw = HashMap<String, Element>()
        private var defaultParagraph: String? = null
        private val docDefaults: Fmt
        private val paraDefaults: ParaFmt
        private val paragraphCache = HashMap<String, StyleInfo>()
        private val characterCache = HashMap<String, Fmt>()

        init {
            val root = doc?.documentElement
            if (root != null) {
                for (s in elementChildren(root).filter { it.localName == "style" }) {
                    val id = s.getAttributeNS(W, "styleId")
                    raw[id] = s
                    val isDefault = s.getAttributeNS(W, "default").let { it == "1" || it == "true" }
                    if (s.getAttributeNS(W, "type") == "paragraph" && isDefault) defaultParagraph = id
                }
            }
            docDefaults = root?.let { child(it, "docDefaults") }?.let { child(it, "rPrDefault") }
                ?.let { child(it, "rPr") }.let { Fmt.of(it) }
            paraDefaults = root?.let { child(it, "docDefaults") }?.let { child(it, "pPrDefault") }
                ?.let { child(it, "pPr") }.let { ParaFmt.of(it) }
        }

        /** The id of the paragraph style named [name] (case-insensitive), or null. */
        fun idByName(name: String): String? = raw.entries.firstOrNull { (_, s) ->
            s.getAttributeNS(W, "type") == "paragraph" &&
                child(s, "name")?.getAttributeNS(W, "val")?.equals(name, ignoreCase = true) == true
        }?.key

        fun paragraph(id: String): StyleInfo {
            val key = id.ifEmpty { defaultParagraph ?: "" }
            return paragraphCache.getOrPut(key) { resolveParagraph(key, HashSet()) }
        }

        private fun resolveParagraph(id: String, seen: MutableSet<String>): StyleInfo {
            val s = raw[id]
            if (s == null || !seen.add(id)) return StyleInfo(run = docDefaults, para = paraDefaults)
            val parent = s.getAttributeNS(W, "basedOn").takeIf { it.isNotEmpty() }
                ?.let { resolveParagraph(it, seen) } ?: StyleInfo(run = docDefaults, para = paraDefaults)
            val pPr = child(s, "pPr")
            val numPr = pPr?.let { child(it, "numPr") }
            return StyleInfo(
                name = child(s, "name")?.getAttributeNS(W, "val") ?: id,
                outline = pPr?.let { child(it, "outlineLvl") }?.getAttributeNS(W, "val")?.toIntOrNull()
                    ?: headingLevelFromName(child(s, "name")?.getAttributeNS(W, "val"))
                    ?: parent.outline,
                jc = pPr?.let { child(it, "jc") }?.getAttributeNS(W, "val") ?: parent.jc,
                numId = numPr?.let { child(it, "numId") }?.getAttributeNS(W, "val")?.toIntOrNull() ?: parent.numId,
                ilvl = numPr?.let { child(it, "ilvl") }?.getAttributeNS(W, "val")?.toIntOrNull() ?: parent.ilvl,
                indent = pPr?.let { child(it, "ind") }?.let { twips(it) } ?: parent.indent,
                run = parent.run.merge(Fmt.of(child(s, "rPr"))),
                para = parent.para.merge(ParaFmt.of(pPr)),
            )
        }

        fun character(id: String?): Fmt {
            if (id.isNullOrEmpty()) return Fmt()
            return characterCache.getOrPut(id) { resolveCharacter(id, HashSet()) }
        }

        private fun resolveCharacter(id: String, seen: MutableSet<String>): Fmt {
            val s = raw[id]
            if (s == null || !seen.add(id)) return Fmt()
            val parent = s.getAttributeNS(W, "basedOn").takeIf { it.isNotEmpty() }
                ?.let { resolveCharacter(it, seen) } ?: Fmt()
            return parent.merge(Fmt.of(child(s, "rPr")))
        }

        /** Built-in heading styles whose definitions omit outlineLvl ("heading 1" … "heading 9"). */
        private fun headingLevelFromName(name: String?): Int? =
            name?.lowercase()?.let { Regex("^heading ([1-9])$").find(it) }?.groupValues?.get(1)?.toInt()?.minus(1)
    }

    // ---------------------------------------------------------------- numbering

    /** numbering.xml: list labels, counted in document order per list (numId). */
    private class Numbering(doc: Document?) {
        private class Level(val fmt: String, val text: String, val start: Int, val indent: Int)

        private val abstracts = HashMap<String, Array<Level?>>()
        private val nums = HashMap<Int, Pair<String, Map<Int, Int>>>() // numId → abstractId, start overrides
        private val counters = HashMap<Int, IntArray>()

        init {
            val root = doc?.documentElement
            if (root != null) {
                for (a in elementChildren(root).filter { it.localName == "abstractNum" }) {
                    val levels = arrayOfNulls<Level>(9)
                    for (l in elementChildren(a).filter { it.localName == "lvl" }) {
                        val i = l.getAttributeNS(W, "ilvl").toIntOrNull() ?: continue
                        if (i !in 0..8) continue
                        levels[i] = Level(
                            fmt = child(l, "numFmt")?.getAttributeNS(W, "val") ?: "decimal",
                            text = child(l, "lvlText")?.getAttributeNS(W, "val") ?: "%${i + 1}.",
                            start = child(l, "start")?.getAttributeNS(W, "val")?.toIntOrNull() ?: 1,
                            indent = child(l, "pPr")?.let { child(it, "ind") }?.let { twips(it) } ?: (720 * (i + 1)),
                        )
                    }
                    abstracts[a.getAttributeNS(W, "abstractNumId")] = levels
                }
                for (n in elementChildren(root).filter { it.localName == "num" }) {
                    val id = n.getAttributeNS(W, "numId").toIntOrNull() ?: continue
                    val abstractId = child(n, "abstractNumId")?.getAttributeNS(W, "val") ?: continue
                    val overrides = HashMap<Int, Int>()
                    for (o in elementChildren(n).filter { it.localName == "lvlOverride" }) {
                        val i = o.getAttributeNS(W, "ilvl").toIntOrNull() ?: continue
                        child(o, "startOverride")?.getAttributeNS(W, "val")?.toIntOrNull()?.let { overrides[i] = it }
                    }
                    nums[id] = abstractId to overrides
                }
            }
        }

        private fun levels(numId: Int): Pair<Array<Level?>, Map<Int, Int>>? {
            val (abstractId, overrides) = nums[numId] ?: return null
            return (abstracts[abstractId] ?: return null) to overrides
        }

        fun indent(numId: Int, ilvl: Int): Int? = levels(numId)?.first?.getOrNull(ilvl)?.indent

        fun definition(numId: Int): ListDef? {
            val (levels, overrides) = levels(numId) ?: return null
            return ListDef(levels.map { l -> l?.let { ListLevel(it.fmt, it.text, it.start) } }, overrides)
        }

        fun label(numId: Int, ilvl: Int): String? {
            val (levels, overrides) = levels(numId) ?: return null
            val i = ilvl.coerceIn(0, 8)
            val level = levels[i] ?: return null
            val count = counters.getOrPut(numId) { IntArray(9) { -1 } }
            count[i] = if (count[i] < 0) (overrides[i] ?: level.start) else count[i] + 1
            for (deeper in i + 1 until 9) count[deeper] = -1
            if (level.fmt == "none") return ""
            if (level.fmt == "bullet") return BULLETS[i % BULLETS.size]
            return Regex("%([1-9])").replace(level.text) { m ->
                val k = m.groupValues[1].toInt() - 1
                val l = levels.getOrNull(k)
                val n = if (count[k] < 0) (overrides[k] ?: l?.start ?: 1) else count[k]
                format(n, l?.fmt ?: "decimal")
            }
        }

        companion object {
            val BULLETS = listOf("•", "◦", "▪")

            fun format(n: Int, fmt: String): String = when (fmt) {
                "lowerLetter" -> letters(n).lowercase()
                "upperLetter" -> letters(n)
                "lowerRoman" -> roman(n).lowercase()
                "upperRoman" -> roman(n)
                "decimalZero" -> n.toString().padStart(2, '0')
                else -> n.toString()
            }

            private fun letters(n: Int): String {
                if (n <= 0) return n.toString()
                val ch = 'A' + (n - 1) % 26
                return ch.toString().repeat((n - 1) / 26 + 1)
            }

            private fun roman(n: Int): String {
                if (n <= 0 || n >= 4000) return n.toString()
                val table = listOf(1000 to "M", 900 to "CM", 500 to "D", 400 to "CD", 100 to "C", 90 to "XC",
                    50 to "L", 40 to "XL", 10 to "X", 9 to "IX", 5 to "V", 4 to "IV", 1 to "I")
                var rest = n
                return buildString {
                    for ((v, s) in table) while (rest >= v) {
                        append(s)
                        rest -= v
                    }
                }
            }
        }
    }

    private fun pageSetup(sect: Element?): PageSetup? {
        if (sect == null) return null
        val sz = child(sect, "pgSz")
        val mar = child(sect, "pgMar")
        fun n(e: Element?, a: String, d: Int) = e?.getAttributeNS(W, a)?.toIntOrNull() ?: d
        return PageSetup(
            width = n(sz, "w", 12240), height = n(sz, "h", 15840),
            top = n(mar, "top", 1440), right = n(mar, "right", 1440), bottom = n(mar, "bottom", 1440), left = n(mar, "left", 1440),
            landscape = sz?.getAttributeNS(W, "orient") == "landscape",
        )
    }

    // ---------------------------------------------------------------- XML / zip helpers

    private fun twips(ind: Element): Int? =
        (ind.getAttributeNS(W, "left").ifEmpty { ind.getAttributeNS(W, "start") }).toIntOrNull()

    private fun preview(el: Element): String =
        el.textContent.replace(Regex("\\s+"), " ").trim().take(120)

    fun isOn(el: Element?): Boolean {
        if (el == null) return false
        val v = el.getAttributeNS(W, "val")
        return v.isEmpty() || (v != "0" && v != "false" && v != "off")
    }

    fun child(parent: Element, localName: String): Element? =
        elementChildren(parent).firstOrNull { it.localName == localName && it.namespaceURI == W }

    fun elementChildren(parent: Element): List<Element> {
        val out = ArrayList<Element>()
        var n: Node? = parent.firstChild
        while (n != null) {
            if (n is Element) out.add(n)
            n = n.nextSibling
        }
        return out
    }

    fun parse(xml: ByteArray): Document {
        val f = DocumentBuilderFactory.newInstance()
        f.isNamespaceAware = true
        f.isExpandEntityReferences = false
        // Untrusted input: no DTDs, no external entities (best effort per parser; DOCTYPE refused).
        runCatching { f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true) }
        runCatching { f.setFeature("http://xml.org/sax/features/external-general-entities", false) }
        runCatching { f.setFeature("http://xml.org/sax/features/external-parameter-entities", false) }
        val head = String(xml, 0, minOf(xml.size, 2048), Charsets.UTF_8)
        require(!head.contains("<!DOCTYPE", ignoreCase = true)) { "DOCTYPE refused" }
        return f.newDocumentBuilder().parse(ByteArrayInputStream(xml))
    }

    fun readEntry(zip: ZipFile, entry: ZipEntry): ByteArray {
        zip.getInputStream(entry).use { input ->
            val out = ByteArrayOutputStream()
            val buf = ByteArray(64 * 1024)
            var total = 0L
            while (true) {
                val n = input.read(buf)
                if (n < 0) break
                total += n
                require(total <= MAX_PART_BYTES) { "${entry.name} expands past $MAX_PART_BYTES bytes" }
                out.write(buf, 0, n)
            }
            return out.toByteArray()
        }
    }

    fun checkEntryName(name: String) {
        require(!name.startsWith("/") && !name.split('/', '\\').contains("..")) { "unsafe entry name: $name" }
    }
}
