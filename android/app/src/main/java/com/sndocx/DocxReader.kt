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

    data class Result(val blocks: List<Block>, val report: Report)

    // ---------------------------------------------------------------- entry points

    fun read(file: File): Result {
        require(file.isFile) { "not a file: ${file.path}" }
        require(file.length() <= MAX_DOCX_BYTES) { "too large: ${file.length()} bytes" }
        ZipFile(file).use { zip ->
            fun part(name: String): Document? = zip.getEntry(name)?.let { parse(readEntry(zip, it)) }
            for (e in zip.entries()) checkEntryName(e.name)
            val document = part(DOCUMENT_PART) ?: throw IllegalStateException("no $DOCUMENT_PART in package")
            return read(document, part("word/styles.xml"), part("word/numbering.xml"))
        }
    }

    fun read(document: Document, styles: Document?, numbering: Document?): Result {
        val ctx = Context(Styles(styles), Numbering(numbering))
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
        return Result(blocks, ctx.report())
    }

    // ---------------------------------------------------------------- paragraphs

    private class Context(val styles: Styles, val numbering: Numbering) {
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
            collectRuns(p, style.run, link = false, out = runs)
            return Paragraph(
                index = index,
                styleId = styleId,
                kind = kind,
                level = if (kind == "heading") outline!! + 1 else 0,
                align = align,
                indentTwips = indent,
                listLabel = label,
                runs = mergeAdjacent(runs),
            )
        }

        /**
         * Runs in reading order. Descends through containers the reader shows inline
         * (hyperlinks, insertions, fields, smart tags, inline content controls) and skips
         * deleted or moved-away text, which Word does not show either.
         */
        private fun collectRuns(parent: Element, base: Fmt, link: Boolean, out: MutableList<Run>) {
            for (el in elementChildren(parent)) {
                if (el.namespaceURI != W) {
                    // mc:AlternateContent and other foreign wrappers: an object the reader can't show.
                    if (el.localName == "AlternateContent") {
                        images++
                        out.add(Run(OBJECT.toString(), obj = "object"))
                    }
                    continue
                }
                when (el.localName) {
                    "r" -> run(el, base, link, out)
                    "hyperlink" -> collectRuns(el, base, true, out)
                    "ins", "moveTo" -> {
                        tracked++
                        collectRuns(el, base, link, out)
                    }
                    "del", "moveFrom" -> tracked++
                    "fldSimple" -> {
                        fields++
                        collectRuns(el, base, link, out)
                    }
                    "smartTag", "customXml" -> collectRuns(el, base, link, out)
                    "sdt" -> {
                        contentControls++
                        child(el, "sdtContent")?.let { collectRuns(it, base, link, out) }
                    }
                    "commentRangeStart" -> comments++
                }
            }
        }

        private fun run(r: Element, base: Fmt, link: Boolean, out: MutableList<Run>) {
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
    ) {
        fun merge(over: Fmt) = Fmt(
            over.bold ?: bold,
            over.italic ?: italic,
            over.underline ?: underline,
            over.strike ?: strike,
            over.highlight ?: highlight,
            over.superscript ?: superscript,
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
    )

    /** styles.xml, resolved through basedOn chains (cycle-safe). */
    private class Styles(doc: Document?) {
        private val raw = HashMap<String, Element>()
        private var defaultParagraph: String? = null
        private val docDefaults: Fmt
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
        }

        fun paragraph(id: String): StyleInfo {
            val key = id.ifEmpty { defaultParagraph ?: "" }
            return paragraphCache.getOrPut(key) { resolveParagraph(key, HashSet()) }
        }

        private fun resolveParagraph(id: String, seen: MutableSet<String>): StyleInfo {
            val s = raw[id]
            if (s == null || !seen.add(id)) return StyleInfo(run = docDefaults)
            val parent = s.getAttributeNS(W, "basedOn").takeIf { it.isNotEmpty() }
                ?.let { resolveParagraph(it, seen) } ?: StyleInfo(run = docDefaults)
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
