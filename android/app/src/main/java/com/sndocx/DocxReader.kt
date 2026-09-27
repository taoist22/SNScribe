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
    private const val DOCUMENT_PART = "word/document.xml"
    const val DOCUMENT_RELS = "word/_rels/document.xml.rels"
    const val REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"

    /** The theme part (for theme fonts), through the document's relationships. */
    fun themePart(rels: Document?): String = relTarget(rels, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        ?: "word/theme/theme1.xml"

    private fun relTarget(rels: Document?, type: String): String? {
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
        /** For a U+FFFC run: "image", "note" or "object". */
        val obj: String? = null,
        /** Inside a field or content control: shown and formattable, but its text is not editable. */
        val locked: Boolean = false,
        /** Font family name, theme fonts resolved; null = not set anywhere. */
        val font: String? = null,
        /** Size in half-points (w:sz); null = not set anywhere. */
        val size: Int? = null,
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

    /** [lists] holds the definitions of every list a paragraph uses, so the screen can recount. */
    data class Result(val blocks: List<Block>, val report: Report, val lists: Map<Int, ListDef> = emptyMap())

    // ---------------------------------------------------------------- entry points

    fun read(file: File): Result {
        require(file.isFile) { "not a file: ${file.path}" }
        require(file.length() <= MAX_DOCX_BYTES) { "too large: ${file.length()} bytes" }
        ZipFile(file).use { zip ->
            fun part(name: String): Document? = zip.getEntry(name)?.let { parse(readEntry(zip, it)) }
            for (e in zip.entries()) checkEntryName(e.name)
            val document = part(DOCUMENT_PART) ?: throw IllegalStateException("no $DOCUMENT_PART in package")
            val rels = part(DOCUMENT_RELS)
            return read(document, part("word/styles.xml"), part(numberingPart(rels)), part(themePart(rels)))
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
        return Result(blocks, ctx.report(), used.mapNotNull { id -> ctx.numbering.definition(id)?.let { id to it } }.toMap())
    }

    // ---------------------------------------------------------------- paragraphs

    /** The theme's heading (major) and body (minor) Latin fonts. */
    private data class ThemeFonts(val major: String?, val minor: String?)

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
            )
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
                if (seg.isRun) run(seg.el, base, seg.link, seg.locked, out) else out.add(Run(OBJECT.toString(), obj = "object", locked = true))
            }
        }

        private fun run(r: Element, base: Fmt, link: Boolean, locked: Boolean, out: MutableList<Run>) {
            val rPr = child(r, "rPr")
            val charStyle = rPr?.let { child(it, "rStyle") }?.getAttributeNS(W, "val")
            val fmt = base.merge(styles.character(charStyle)).merge(Fmt.of(rPr))
            val isLink = link || charStyle.equals("Hyperlink", ignoreCase = true)
            fun add(text: String, obj: String? = null) {
                if (text.isEmpty()) return
                out.add(
                    Run(
                        text = text,
                        bold = fmt.bold == true,
                        italic = fmt.italic == true,
                        underline = fmt.underline == true || isLink,
                        strike = fmt.strike == true,
                        highlight = fmt.highlight == true,
                        link = isLink,
                        superscript = fmt.superscript == true || obj == "note",
                        obj = obj,
                        locked = locked,
                        font = when (fmt.font) {
                            "+major" -> theme.major
                            "+minor" -> theme.minor
                            else -> fmt.font
                        },
                        size = fmt.size,
                    ),
                )
            }
            val text = StringBuilder()
            for (c in elementChildren(r)) {
                val t = textOf(c)
                if (t.isEmpty()) {
                    if (c.localName == "fldChar") fields++
                    continue
                }
                if (t[0] == OBJECT) {
                    add(text.toString())
                    text.clear()
                    val kind = when (c.localName) {
                        "drawing", "pict" -> "image".also { images++ }
                        "footnoteReference", "endnoteReference" -> "note"
                        else -> "object"
                    }
                    add(t, kind)
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
    class Segment(val el: Element, val isRun: Boolean, val link: Boolean, var locked: Boolean = false) {
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
    private fun walk(parent: Element, link: Boolean, out: MutableList<Segment>, onMarker: (Element) -> Unit, locked: Boolean) {
        for (el in elementChildren(parent)) {
            if (el.namespaceURI != W) {
                if (el.localName == "AlternateContent") {
                    onMarker(el)
                    out.add(Segment(el, false, link, locked))
                }
                continue
            }
            when (el.localName) {
                "r" -> out.add(Segment(el, true, link, locked))
                "hyperlink" -> walk(el, true, out, onMarker, locked)
                "ins", "moveTo" -> {
                    onMarker(el)
                    walk(el, link, out, onMarker, locked)
                }
                "fldSimple" -> {
                    onMarker(el)
                    walk(el, link, out, onMarker, true)
                }
                "smartTag", "customXml" -> walk(el, link, out, onMarker, locked)
                "sdt" -> {
                    onMarker(el)
                    child(el, "sdtContent")?.let { walk(it, link, out, onMarker, true) }
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
        val highlight: Boolean? = null,
        val superscript: Boolean? = null,
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
            over.superscript ?: superscript,
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
                    highlight = hl?.let { it.isNotEmpty() && it != "none" },
                    superscript = va?.let { it == "superscript" },
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
