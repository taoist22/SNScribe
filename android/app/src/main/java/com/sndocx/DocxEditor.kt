package com.sndocx

import com.sndocx.DocxReader.W
import com.sndocx.DocxReader.child
import com.sndocx.DocxReader.elementChildren
import org.w3c.dom.Document
import org.w3c.dom.Element
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.FileOutputStream
import java.security.MessageDigest
import java.util.zip.CRC32
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream
import javax.xml.XMLConstants
import javax.xml.transform.OutputKeys
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult

/**
 * Applies formatting edits to a copy of a .docx. Pure Kotlin + javax.xml, like [DocxReader].
 *
 * The original package is the source of truth: edits are applied to its word/document.xml
 * (and word/styles.xml only when a heading style has to be added); every other part is
 * copied with its bytes unchanged. Offsets are into the paragraph text as
 * [DocxReader.segments] defines it — the same text the reader shows — so an edit lands on
 * the characters the user selected.
 *
 * Ops apply in order; each op's offsets refer to the text left by the ops before it (the
 * text the screen showed when the op was made). The saved copy is verified by re-reading
 * it: every paragraph's text must equal the original with the text ops applied as plain
 * string edits — and, when the screen sent them, the texts the screen expects.
 */
object DocxEditor {
    sealed class Op {
        /** prop: "b", "i", "u" or "h". */
        data class Format(val para: Int, val start: Int, val end: Int, val prop: String, val on: Boolean) : Op()

        /** kind: "heading1", "heading2", "title" or "normal". */
        data class Style(val para: Int, val kind: String) : Op()

        /**
         * Replaces characters [start, end) of a paragraph with [text]: delete when [text] is
         * empty, insert when start == end. New text takes the formatting of the first
         * replaced character, or when inserting, of the character before it (Word's rule).
         */
        data class Text(val para: Int, val start: Int, val end: Int, val text: String) : Op()

        /**
         * Splits a paragraph at [offset] (Enter): the text before stays as paragraph [para],
         * the rest becomes paragraph [para] + 1, and later paragraphs move down one.
         */
        data class Split(val para: Int, val offset: Int) : Op()

        /** Joins paragraph [para] onto the end of [para] - 1 (Backspace at its start). */
        data class Join(val para: Int) : Op()

        /**
         * Makes a paragraph a list item, or not. [kind] "number" / "bullet" / "none".
         * [listId]: a number = an existing list of the document (continue it); anything else
         * names a new list made in this save — every op with the same name joins the same
         * new list, which starts at 1.
         */
        data class ListItem(val para: Int, val kind: String, val listId: String) : Op()

        /** Sets the font family and/or the size (half-points) of characters [start, end). */
        data class RunStyle(val para: Int, val start: Int, val end: Int, val font: String?, val size: Int?) : Op()

        /**
         * Paragraph formatting; null = leave as is. [align] left/center/right/justify;
         * [line] in 240ths of a line ([lineRule] "auto") or twips; [before]/[after] in twips;
         * [first] first-line indent in twips, negative = hanging, 0 = none;
         * [pageBreakBefore] starts the paragraph on a new page.
         */
        data class ParaProps(
            val para: Int,
            val align: String? = null,
            val line: Int? = null,
            val lineRule: String? = null,
            val before: Int? = null,
            val after: Int? = null,
            val first: Int? = null,
            val pageBreakBefore: Boolean? = null,
        ) : Op()

        /** Page size, orientation and margins (twips) for every section; null = unchanged. */
        data class PageSetup(
            val width: Int? = null,
            val height: Int? = null,
            val landscape: Boolean? = null,
            val top: Int? = null,
            val right: Int? = null,
            val bottom: Int? = null,
            val left: Int? = null,
        ) : Op()

        /** The document's default font and size (half-points): what text without its own takes. */
        data class Defaults(val font: String?, val size: Int?) : Op()
    }

    data class Saved(val dest: File, val changedParts: List<String>, val notes: List<String>)

    private const val DOCUMENT_PART = "word/document.xml"
    private const val STYLES_PART = "word/styles.xml"
    private const val CONTENT_TYPES = "[Content_Types].xml"
    private const val NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    private const val PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
    private const val PKG_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

    /** Schema order of w:pPr children (CT_PPrBase + rPr, sectPr, pPrChange). */
    private val PPR_ORDER = listOf(
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
        "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
        "sectPr", "pPrChange",
    )

    /** Schema order of w:rPr children (CT_RPr is a sequence; Word calls a file damaged otherwise). */
    private val RPR_ORDER = listOf(
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
        "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath",
    )

    // ---------------------------------------------------------------- save

    /**
     * Writes [src] with [ops] applied to [dest]. The package is built and verified in
     * [workDir] (private storage, where a failed attempt can be deleted) and only then
     * written to [dest], so [dest] never holds an unverified file. Throws with the reason
     * when anything fails; [dest] is then untouched.
     */
    fun save(src: File, ops: List<Op>, dest: File, workDir: File, expected: List<String> = emptyList()): Saved {
        require(src.length() <= DocxReader.MAX_DOCX_BYTES) { "too large: ${src.length()} bytes" }
        val notes = ArrayList<String>()
        val changed = ArrayList<String>()
        val added = LinkedHashMap<String, ByteArray>()
        workDir.mkdirs()
        val temp = File(workDir, "saving-${System.nanoTime()}.docx")
        try {
            ZipFile(src).use { zip ->
                fun part(name: String) = zip.getEntry(name)?.let { DocxReader.parse(DocxReader.readEntry(zip, it)) }
                val document = part(DOCUMENT_PART) ?: error("no $DOCUMENT_PART")
                val styles = part(STYLES_PART)
                val rels = part(DocxReader.DOCUMENT_RELS)
                val numberingPath = DocxReader.numberingPart(rels)
                val numbering = part(numberingPath)
                val lists = Lists(numbering)
                val stylesChanged = apply(document, styles, ops, notes, lists)

                val replaced = LinkedHashMap<String, ByteArray>()
                if (ops.isNotEmpty()) replaced[DOCUMENT_PART] = serialize(document)
                if (stylesChanged && styles != null) replaced[STYLES_PART] = serialize(styles)
                if (lists.changed) {
                    val bytes = serialize(lists.doc)
                    if (numbering != null) {
                        replaced[numberingPath] = bytes
                    } else {
                        // A document with no lists yet: add the part and register it.
                        check(zip.getEntry(numberingPath) == null) { "unreadable $numberingPath" }
                        added[numberingPath] = bytes
                        val relsDoc = rels ?: error("no ${DocxReader.DOCUMENT_RELS}")
                        addRelationship(relsDoc, DocxReader.REL_NUMBERING, numberingPath.removePrefix("word/"))
                        replaced[DocxReader.DOCUMENT_RELS] = serialize(relsDoc)
                        val types = part(CONTENT_TYPES) ?: error("no $CONTENT_TYPES")
                        addOverride(types, "/$numberingPath", NUMBERING_TYPE)
                        replaced[CONTENT_TYPES] = serialize(types)
                    }
                }

                ZipOutputStream(FileOutputStream(temp)).use { out ->
                    for (entry in zip.entries()) {
                        DocxReader.checkEntryName(entry.name)
                        putEntry(out, entry, replaced[entry.name] ?: DocxReader.readEntry(zip, entry))
                    }
                    for ((name, bytes) in added) putEntry(out, ZipEntry(name).apply { method = ZipEntry.DEFLATED }, bytes)
                }
                changed.addAll(replaced.keys)
                changed.addAll(added.keys)
            }
            verify(src, temp, changed, ops, expected, added.keys)
            copy(temp, dest)
            return Saved(dest, changed, notes)
        } finally {
            temp.delete()
        }
    }

    /**
     * Every paragraph's text is exactly what the ops say (original + text edits as string
     * splices) and what the screen expected; every untouched part is byte-identical.
     */
    private fun verify(src: File, written: File, changed: List<String>, ops: List<Op>, expected: List<String>, added: Set<String>) {
        val before = DocxReader.read(src).blocks.filterIsInstance<DocxReader.Paragraph>()
        val after = DocxReader.read(written).blocks.filterIsInstance<DocxReader.Paragraph>()
        // The same edits on plain strings: what every paragraph must now say.
        val want = before.map { it.text }.toMutableList()
        for (op in ops) when (op) {
            is Op.Text -> want[op.para] = want[op.para].let { it.substring(0, op.start) + op.text + it.substring(op.end) }
            is Op.Split -> want[op.para].let {
                want[op.para] = it.substring(0, op.offset)
                want.add(op.para + 1, it.substring(op.offset))
            }
            is Op.Join -> {
                want[op.para - 1] = want[op.para - 1] + want[op.para]
                want.removeAt(op.para)
            }
            else -> {}
        }
        check(after.size == want.size) { "verify: ${after.size} paragraphs, the edits describe ${want.size}" }
        after.forEachIndexed { i, p ->
            check(p.text == want[i]) { "verify: paragraph $i does not have the text the edits describe" }
        }
        if (expected.isNotEmpty()) {
            check(expected.size == after.size) { "verify: the screen showed ${expected.size} paragraphs, the file has ${after.size}" }
            after.forEachIndexed { i, p -> check(p.text == expected[i]) { "verify: paragraph $i differs from what the screen showed" } }
        }
        val ha = hashes(src)
        val hb = hashes(written)
        check(hb.keys == ha.keys + added) { "verify: package parts differ: ${ha.keys - hb.keys} / ${hb.keys - ha.keys}" }
        for ((name, hash) in ha) {
            if (name !in changed) check(hb[name] == hash) { "verify: $name changed but was not edited" }
        }
    }

    // ---------------------------------------------------------------- apply

    /** Applies [ops] to the DOM. Returns whether styles.xml was changed (a style was added). */
    fun apply(document: Document, styles: Document?, ops: List<Op>, notes: MutableList<String>, lists: Lists = Lists(null)): Boolean {
        var paragraphs = bodyParagraphs(document)
        val styleIds = StyleIds(styles)
        var stylesTouched = false
        for (op in ops) {
            // Document-wide edits address no paragraph.
            if (op is Op.PageSetup) {
                pageSetup(document, op, notes)
                continue
            }
            if (op is Op.Defaults) {
                if (styles != null) {
                    defaults(styles, op, notes)
                    stylesTouched = true
                }
                continue
            }
            val index = when (op) {
                is Op.Format -> op.para
                is Op.Style -> op.para
                is Op.Text -> op.para
                is Op.Split -> op.para
                is Op.Join -> op.para
                is Op.ListItem -> op.para
                is Op.RunStyle -> op.para
                is Op.ParaProps -> op.para
                is Op.PageSetup, is Op.Defaults -> error("unreachable")
            }
            val p = checkNotNull(paragraphs.getOrNull(index)) { "no paragraph $index for $op" }
            when (op) {
                is Op.Format -> format(document, p, op, notes)
                is Op.Style -> setStyle(document, p, styleIds.idFor(op.kind))
                is Op.Text -> replaceText(document, p, op, notes)
                is Op.Split -> splitParagraph(document, p, op, styleIds, notes)
                is Op.Join -> joinParagraph(checkNotNull(paragraphs.getOrNull(index - 1)) { "nothing before paragraph $index to join onto" }, p, op, notes)
                is Op.ListItem -> setListItem(document, p, op, lists, styleIds, notes)
                is Op.RunStyle -> runStyle(document, p, op, notes)
                is Op.ParaProps -> paraProps(document, p, op, notes)
                is Op.PageSetup, is Op.Defaults -> {}
            }
            // Splits and joins renumber the paragraphs after them.
            if (op is Op.Split || op is Op.Join) paragraphs = bodyParagraphs(document)
        }
        return styleIds.added || stylesTouched
    }

    fun bodyParagraphs(document: Document): List<Element> {
        val body = child(document.documentElement, "body") ?: return emptyList()
        return elementChildren(body).filter { it.localName == "p" && it.namespaceURI == W }
    }

    private fun format(document: Document, p: Element, op: Op.Format, notes: MutableList<String>) {
        if (op.end <= op.start) return
        splitAt(p, op.start)
        splitAt(p, op.end)
        var offset = 0
        var touched = 0
        for (seg in DocxReader.segments(p)) {
            val len = seg.length
            if (seg.isRun && len > 0 && offset >= op.start && offset + len <= op.end) {
                setRunProperty(document, seg.el, op.prop, op.on)
                touched++
            }
            offset += len
        }
        notes.add("${op.prop}=${op.on} p${op.para} [${op.start},${op.end}): $touched run(s)")
    }

    /**
     * Replaces [op.start, op.end) of paragraph [p] with [op.text] (see [Op.Text]).
     * Refuses — throws — when the range touches an image, note reference, other object, or
     * field / content-control text: deleting those would orphan or corrupt structure the
     * editor does not manage yet.
     */
    private fun replaceText(document: Document, p: Element, op: Op.Text, notes: MutableList<String>) {
        val segs = DocxReader.segments(p)
        val total = segs.sumOf { it.length }
        require(op.start in 0..op.end && op.end <= total) { "text edit [${op.start},${op.end}) outside paragraph ${op.para} (length $total)" }

        // What the range covers must be plain, unlocked text.
        var offset = 0
        for (seg in segs) {
            val s = offset
            offset += seg.length
            if (seg.length > 0 && s < op.end && offset > op.start) {
                check(seg.isRun && !seg.locked && !hasObject(seg.el, s, op.start, op.end)) {
                    "paragraph ${op.para}: the edit touches an image, note marker, field or content control"
                }
            }
        }

        // Formatting source: first replaced character, else the one before, else the one after.
        val source = sourceRun(segs, if (op.end > op.start) op.start else op.start - 1) ?: sourceRun(segs, op.start)
        check(source == null || !source.locked) { "paragraph ${op.para}: can't type inside a field or content control" }
        val rPr = source?.el?.let { child(it, "rPr") }?.cloneNode(true) as Element?
        rPr?.let { r -> elementChildren(r).filter { it.localName == "rPrChange" }.forEach { r.removeChild(it) } }

        splitAt(p, op.start)
        splitAt(p, op.end)
        offset = 0
        for (seg in DocxReader.segments(p)) {
            val s = offset
            offset += seg.length
            if (seg.length > 0 && s >= op.start && offset <= op.end) {
                val parent = seg.el.parentNode
                parent.removeChild(seg.el)
                // A hyperlink (or other container) emptied by the deletion goes too.
                if (parent is Element && parent !== p && elementChildren(parent).none { it.localName == "r" }) {
                    parent.parentNode?.removeChild(parent)
                }
            }
        }

        if (op.text.isNotEmpty()) {
            val run = document.createElementNS(W, "w:r")
            rPr?.let { run.appendChild(it) }
            val parts = op.text.split('\t')
            parts.forEachIndexed { i, part ->
                if (i > 0) run.appendChild(document.createElementNS(W, "w:tab"))
                if (part.isNotEmpty()) {
                    val t = document.createElementNS(W, "w:t")
                    t.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
                    t.textContent = part
                    run.appendChild(t)
                }
            }
            // Between the piece ending at start and the piece starting there: after the former
            // when there is one (so typing at the end of a link extends it, as the screen shows).
            var before: Element? = null
            var after: Element? = null
            offset = 0
            for (seg in DocxReader.segments(p)) {
                val s = offset
                offset += seg.length
                if (seg.length > 0 && offset == op.start) before = seg.el
                if (seg.length > 0 && s == op.start && after == null) after = seg.el
            }
            when {
                before != null -> before.parentNode.insertBefore(run, before.nextSibling)
                after != null -> after.parentNode.insertBefore(run, after)
                else -> p.appendChild(run)
            }
        }
        notes.add("text p${op.para} [${op.start},${op.end}) → ${op.text.length} chars")
    }

    /**
     * Enter: splits [p] at [op.offset]. A new paragraph holding everything before the offset
     * is inserted in front of [p], with a copy of [p]'s properties; [p] itself keeps the rest
     * and its paragraph mark — so a section break (w:sectPr) stays at the end, where Word
     * keeps it. Links and other containers across the offset are divided in two.
     * Refused inside a field or content control.
     */
    private fun splitParagraph(document: Document, p: Element, op: Op.Split, styleIds: StyleIds, notes: MutableList<String>) {
        val segs = DocxReader.segments(p)
        val total = segs.sumOf { it.length }
        require(op.offset in 0..total) { "split at ${op.offset} outside paragraph ${op.para} (length $total)" }
        check(!insideLocked(segs, op.offset)) { "paragraph ${op.para}: can't start a new paragraph inside a field or content control" }
        val styleBefore = child(p, "pPr")?.let { child(it, "pStyle") }?.getAttributeNS(W, "val")

        splitAt(p, op.offset)
        val first = document.createElementNS(W, "w:p")
        child(p, "pPr")?.let { pPr ->
            val copy = pPr.cloneNode(true) as Element
            child(copy, "sectPr")?.let { copy.removeChild(it) }
            first.appendChild(copy)
        }
        moveBefore(p, first, op.offset)
        p.parentNode.insertBefore(first, p)

        // Enter at the end of a heading: the new paragraph takes the style's "next" style.
        if (op.offset == total && styleBefore != null) {
            val next = styleIds.nextOf(styleBefore)
            if (next != null && next != styleBefore) setStyle(document, p, next.takeUnless { styleIds.isDefault(it) })
        }
        notes.add("split p${op.para} at ${op.offset}")
    }

    /**
     * Moves the content of [from] that lies before character [cut] into [to], in order.
     * A child straddling [cut] (after splitAt, only a container such as a hyperlink can) is
     * divided: its first part, a shallow copy, goes to [to]. Zero-length markers before
     * [cut] (bookmarks, comment starts) move with the text before them.
     */
    private fun moveBefore(from: Element, to: Element, cut: Int) {
        var pos = 0
        for (c in elementChildren(from)) {
            if (c.localName == "pPr" && c.namespaceURI == W) continue
            if (pos >= cut) break
            val len = DocxReader.lengthOf(c)
            if (pos + len <= cut) {
                from.removeChild(c)
                to.appendChild(c)
                pos += len
                continue
            }
            val piece = c.cloneNode(false) as Element
            if (c.localName == "sdt") {
                // Guarded by the locked check; kept for safety: an sdt is never divided.
                error("can't divide a content control")
            }
            moveBefore(c, piece, cut - pos)
            to.appendChild(piece)
            break
        }
    }

    /** Whether [offset] lies inside field or content-control text (so Enter there is refused). */
    private fun insideLocked(segs: List<DocxReader.Segment>, offset: Int): Boolean {
        var pos = 0
        var before: DocxReader.Segment? = null
        var after: DocxReader.Segment? = null
        for (seg in segs) {
            val s = pos
            pos += seg.length
            if (seg.length == 0) continue
            if (seg.locked && offset > s && offset < pos) return true
            if (pos == offset) before = seg
            if (s == offset && after == null) after = seg
        }
        return before?.locked == true && after?.locked == true
    }

    /**
     * Backspace at the start of [p]: its content moves onto the end of [prev], which keeps
     * its own properties, and [p] is removed. Refused when a table or other block sits
     * between them, or when [prev] ends a section (they are in different sections). A
     * section break on [p] moves to the joined paragraph.
     */
    private fun joinParagraph(prev: Element, p: Element, op: Op.Join, notes: MutableList<String>) {
        val between = ArrayList<Element>()
        var n = prev.nextSibling
        while (n != null && n !== p) {
            if (n is Element) {
                check(n.localName in BODY_MARKERS) { "paragraph ${op.para}: a table or other block lies before it; can't join" }
                between.add(n)
            }
            n = n.nextSibling
        }
        checkNotNull(n) { "paragraph ${op.para} is not after the one it joins" }
        check(child(prev, "pPr")?.let { child(it, "sectPr") } == null) {
            "paragraph ${op.para}: a section break separates these paragraphs; can't join"
        }
        for (el in between) prev.appendChild(el)
        for (c in elementChildren(p)) {
            if (c.localName == "pPr" && c.namespaceURI == W) continue
            prev.appendChild(c)
        }
        // If [p] ended a section, both paragraphs were in that section: the joined paragraph
        // now ends it. sectPr goes last in pPr, before any pPrChange.
        child(p, "pPr")?.let { child(it, "sectPr") }?.let { sect ->
            val pPr = child(prev, "pPr") ?: prev.ownerDocument.createElementNS(W, "w:pPr").also { prev.insertBefore(it, prev.firstChild) }
            pPr.insertBefore(sect, child(pPr, "pPrChange"))
        }
        p.parentNode.removeChild(p)
        notes.add("join p${op.para} onto p${op.para - 1}")
    }

    /** Body-level elements that carry no content and may sit between two joined paragraphs. */
    private val BODY_MARKERS = setOf(
        "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "proofErr",
        "permStart", "permEnd", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    )

    /**
     * Puts [p] in a list, or takes it out. Out: the paragraph's own numbering goes; if its
     * style numbers it, an explicit numId 0 switches that off.
     */
    private fun setListItem(document: Document, p: Element, op: Op.ListItem, lists: Lists, styleIds: StyleIds, notes: MutableList<String>) {
        var pPr = child(p, "pPr")
        pPr?.let { pp -> child(pp, "numPr")?.let { pp.removeChild(it) } }
        val styleId = pPr?.let { child(it, "pStyle") }?.getAttributeNS(W, "val")
        val numId = when (op.kind) {
            "number", "bullet" -> lists.numIdFor(op.listId, op.kind)
            else -> if (styleId != null && styleIds.numbers(styleId)) 0 else null
        }
        if (numId != null) {
            if (pPr == null) {
                pPr = document.createElementNS(W, "w:pPr")
                p.insertBefore(pPr, p.firstChild)
            }
            val numPr = document.createElementNS(W, "w:numPr")
            val ilvl = document.createElementNS(W, "w:ilvl")
            ilvl.setAttributeNS(W, "w:val", "0")
            val id = document.createElementNS(W, "w:numId")
            id.setAttributeNS(W, "w:val", numId.toString())
            numPr.appendChild(ilvl)
            numPr.appendChild(id)
            insertInOrder(pPr!!, numPr, PPR_ORDER)
        }
        notes.add("list p${op.para}: ${op.kind} ${op.listId} → numId $numId")
    }

    /** Inserts [el] among [parent]'s children at its schema position; unknown children sort last. */
    private fun insertInOrder(parent: Element, el: Element, order: List<String>) {
        val rank = order.indexOf(el.localName)
        val before = elementChildren(parent).firstOrNull { (order.indexOf(it.localName).takeIf { i -> i >= 0 } ?: 1000) > rank }
        parent.insertBefore(el, before)
    }

    /**
     * The document's list definitions (numbering.xml), with new lists added as needed. A new
     * list gets its own w:num pointing at one shared definition per kind (decimal "1.",
     * "a.", "i." … or bullets "•", "◦", "▪"), restarted at 1 with a startOverride.
     */
    class Lists(existing: Document?) {
        val doc: Document = existing ?: newNumbering()
        var changed = false
            private set
        private val abstractFor = HashMap<String, Int>()
        private val numFor = HashMap<String, Int>()

        fun numIdFor(listId: String, kind: String): Int {
            listId.toIntOrNull()?.let { return it }
            return numFor.getOrPut(listId) {
                changed = true
                createNum(abstractFor.getOrPut(kind) { createAbstract(kind) })
            }
        }

        private fun root() = doc.documentElement
        private fun children(name: String) = elementChildren(root()).filter { it.localName == name && it.namespaceURI == W }
        private fun el(tag: String, value: String? = null): Element =
            doc.createElementNS(W, "w:$tag").also { if (value != null) it.setAttributeNS(W, "w:val", value) }

        private fun createAbstract(kind: String): Int {
            val id = (children("abstractNum").mapNotNull { it.getAttributeNS(W, "abstractNumId").toIntOrNull() }.maxOrNull() ?: -1) + 1
            val a = doc.createElementNS(W, "w:abstractNum")
            a.setAttributeNS(W, "w:abstractNumId", id.toString())
            a.appendChild(el("multiLevelType", "hybridMultilevel"))
            val numberFormats = listOf("decimal", "lowerLetter", "lowerRoman")
            val bullets = listOf("\u2022", "\u25E6", "\u25AA")
            for (i in 0..8) {
                val lvl = doc.createElementNS(W, "w:lvl")
                lvl.setAttributeNS(W, "w:ilvl", i.toString())
                lvl.appendChild(el("start", "1"))
                if (kind == "bullet") {
                    lvl.appendChild(el("numFmt", "bullet"))
                    lvl.appendChild(el("lvlText", bullets[i % 3]))
                } else {
                    lvl.appendChild(el("numFmt", numberFormats[i % 3]))
                    lvl.appendChild(el("lvlText", "%${i + 1}."))
                }
                lvl.appendChild(el("lvlJc", "left"))
                val pPr = el("pPr")
                val ind = el("ind")
                ind.setAttributeNS(W, "w:left", (720 * (i + 1)).toString())
                ind.setAttributeNS(W, "w:hanging", "360")
                pPr.appendChild(ind)
                lvl.appendChild(pPr)
                a.appendChild(lvl)
            }
            // Every abstractNum comes before every num.
            root().insertBefore(a, children("num").firstOrNull() ?: children("numIdMacAtCleanup").firstOrNull())
            return id
        }

        private fun createNum(abstractId: Int): Int {
            val id = (children("num").mapNotNull { it.getAttributeNS(W, "numId").toIntOrNull() }.maxOrNull() ?: 0) + 1
            val n = doc.createElementNS(W, "w:num")
            n.setAttributeNS(W, "w:numId", id.toString())
            n.appendChild(el("abstractNumId", abstractId.toString()))
            val override = el("lvlOverride")
            override.setAttributeNS(W, "w:ilvl", "0")
            override.appendChild(el("startOverride", "1"))
            n.appendChild(override)
            root().insertBefore(n, children("numIdMacAtCleanup").firstOrNull())
            return id
        }

        companion object {
            private fun newNumbering(): Document {
                val f = javax.xml.parsers.DocumentBuilderFactory.newInstance()
                f.isNamespaceAware = true
                val d = f.newDocumentBuilder().newDocument()
                d.appendChild(d.createElementNS(W, "w:numbering"))
                return d
            }
        }
    }

    private fun addRelationship(rels: Document, type: String, target: String) {
        val root = rels.documentElement
        val ids = HashSet<String>()
        var n = root.firstChild
        while (n != null) {
            if (n is Element) ids.add(n.getAttribute("Id"))
            n = n.nextSibling
        }
        var i = 1
        while ("rId$i" in ids) i++
        val r = rels.createElementNS(PKG_RELS, "Relationship")
        r.setAttribute("Id", "rId$i")
        r.setAttribute("Type", type)
        r.setAttribute("Target", target)
        root.appendChild(r)
    }

    private fun addOverride(types: Document, partName: String, contentType: String) {
        val o = types.createElementNS(PKG_TYPES, "Override")
        o.setAttribute("PartName", partName)
        o.setAttribute("ContentType", contentType)
        types.documentElement.appendChild(o)
    }

    /** The run segment holding character [ch], or null (out of range or not a text run). */
    private fun sourceRun(segs: List<DocxReader.Segment>, ch: Int): DocxReader.Segment? {
        if (ch < 0) return null
        var offset = 0
        for (seg in segs) {
            val s = offset
            offset += seg.length
            if (ch in s until offset) return seg.takeIf { it.isRun && !hasObject(it.el, s, ch, ch + 1) }
        }
        return null
    }

    /** Whether the run's characters that fall in [from, to) include an object (U+FFFC). */
    private fun hasObject(r: Element, runStart: Int, from: Int, to: Int): Boolean {
        var pos = runStart
        for (c in elementChildren(r)) {
            val t = DocxReader.textOf(c)
            for (ch in t) {
                if (pos in from until to && ch == DocxReader.OBJECT) return true
                pos++
            }
        }
        return false
    }

    /**
     * Makes [at] a run boundary. Every run child has a fixed length (text, tab, break,
     * object = 1, markers = 0), so a run can always be divided at a child boundary; only a
     * w:t is cut inside. Order is preserved and both halves keep the run's properties.
     */
    private fun splitAt(p: Element, at: Int) {
        var offset = 0
        for (seg in DocxReader.segments(p)) {
            val len = seg.length
            if (seg.isRun && at > offset && at < offset + len) {
                val r = seg.el
                val cut = at - offset
                val second = r.cloneNode(false) as Element
                child(r, "rPr")?.let { second.appendChild(it.cloneNode(true)) }
                var pos = 0
                for (c in elementChildren(r)) {
                    if (c.localName == "rPr" && c.namespaceURI == W) continue
                    val l = DocxReader.textOf(c).length
                    if (pos >= cut) {
                        r.removeChild(c)
                        second.appendChild(c)
                    } else if (c.localName == "t" && pos + l > cut) {
                        val text = c.textContent
                        val tail = c.cloneNode(false) as Element
                        c.textContent = text.substring(0, cut - pos)
                        tail.textContent = text.substring(cut - pos)
                        c.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
                        tail.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
                        second.appendChild(tail)
                    }
                    pos += l
                }
                r.parentNode.insertBefore(second, r.nextSibling)
                return
            }
            offset += len
        }
    }

    /** Applies [op] to every w:sectPr (the body's and any section breaks in paragraphs). */
    private fun pageSetup(document: Document, op: Op.PageSetup, notes: MutableList<String>) {
        val sects = document.getElementsByTagNameNS(W, "sectPr")
        val order = listOf("headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar")
        for (i in 0 until sects.length) {
            val sect = sects.item(i) as Element
            fun part(name: String) = child(sect, name) ?: document.createElementNS(W, "w:$name").also { insertInOrder(sect, it, order) }
            if (op.width != null || op.height != null || op.landscape != null) {
                val sz = part("pgSz")
                op.width?.let { sz.setAttributeNS(W, "w:w", it.toString()) }
                op.height?.let { sz.setAttributeNS(W, "w:h", it.toString()) }
                op.landscape?.let { if (it) sz.setAttributeNS(W, "w:orient", "landscape") else sz.removeAttributeNS(W, "orient") }
            }
            if (listOf(op.top, op.right, op.bottom, op.left).any { it != null }) {
                val mar = part("pgMar")
                op.top?.let { mar.setAttributeNS(W, "w:top", it.toString()) }
                op.right?.let { mar.setAttributeNS(W, "w:right", it.toString()) }
                op.bottom?.let { mar.setAttributeNS(W, "w:bottom", it.toString()) }
                op.left?.let { mar.setAttributeNS(W, "w:left", it.toString()) }
                // pgMar requires all of these; fill any a document left out.
                for ((a, d) in listOf("header" to "720", "footer" to "720", "gutter" to "0", "top" to "1440", "right" to "1440", "bottom" to "1440", "left" to "1440")) {
                    if (!mar.hasAttributeNS(W, a)) mar.setAttributeNS(W, "w:$a", d)
                }
            }
        }
        notes.add("page setup on ${sects.length} section(s): $op")
    }

    /** The default run font and size in styles.xml (docDefaults/rPrDefault). */
    private fun defaults(styles: Document, op: Op.Defaults, notes: MutableList<String>) {
        val root = styles.documentElement
        fun sub(parent: Element, name: String, before: Element? = null): Element =
            child(parent, name) ?: styles.createElementNS(W, "w:$name").also { parent.insertBefore(it, before) }
        val docDefaults = sub(root, "docDefaults", root.firstChild as? Element)
        val rPrDefault = sub(docDefaults, "rPrDefault", child(docDefaults, "pPrDefault"))
        val rPr = sub(rPrDefault, "rPr")
        op.font?.let { name ->
            child(rPr, "rFonts")?.let { rPr.removeChild(it) }
            val f = styles.createElementNS(W, "w:rFonts")
            for (slot in listOf("ascii", "hAnsi", "eastAsia", "cs")) f.setAttributeNS(W, "w:$slot", name)
            insertInOrder(rPr, f, RPR_ORDER)
        }
        op.size?.let { size ->
            for (tag in listOf("sz", "szCs")) {
                child(rPr, tag)?.let { rPr.removeChild(it) }
                val e = styles.createElementNS(W, "w:$tag")
                e.setAttributeNS(W, "w:val", size.toString())
                insertInOrder(rPr, e, RPR_ORDER)
            }
        }
        notes.add("defaults: $op")
    }

    private fun ensurePPr(document: Document, p: Element): Element =
        child(p, "pPr") ?: document.createElementNS(W, "w:pPr").also { p.insertBefore(it, p.firstChild) }

    /** The paragraph's own element [name] in its pPr, created in schema order if missing. */
    private fun pPrChild(document: Document, pPr: Element, name: String): Element =
        child(pPr, name) ?: document.createElementNS(W, "w:$name").also { insertInOrder(pPr, it, PPR_ORDER) }

    private fun paraProps(document: Document, p: Element, op: Op.ParaProps, notes: MutableList<String>) {
        val pPr = ensurePPr(document, p)
        op.align?.let { a ->
            pPrChild(document, pPr, "jc").setAttributeNS(W, "w:val", if (a == "justify") "both" else a)
        }
        if (op.line != null || op.before != null || op.after != null) {
            val sp = pPrChild(document, pPr, "spacing")
            op.before?.let {
                sp.setAttributeNS(W, "w:before", it.toString())
                sp.removeAttributeNS(W, "beforeAutospacing") // would override the number
            }
            op.after?.let {
                sp.setAttributeNS(W, "w:after", it.toString())
                sp.removeAttributeNS(W, "afterAutospacing")
            }
            op.line?.let {
                sp.setAttributeNS(W, "w:line", it.toString())
                sp.setAttributeNS(W, "w:lineRule", op.lineRule ?: "auto")
            }
        }
        op.first?.let { f ->
            val ind = pPrChild(document, pPr, "ind")
            val oldHanging = ind.getAttributeNS(W, "hanging").toIntOrNull()
            for (a in listOf("firstLine", "hanging", "firstLineChars", "hangingChars")) ind.removeAttributeNS(W, a)
            val leftAttr = if (ind.hasAttributeNS(W, "start") && !ind.hasAttributeNS(W, "left")) "start" else "left"
            val left = ind.getAttributeNS(W, leftAttr).toIntOrNull() ?: 0
            when {
                f > 0 -> ind.setAttributeNS(W, "w:firstLine", f.toString())
                f < 0 -> {
                    ind.setAttributeNS(W, "w:hanging", (-f).toString())
                    // A hanging indent hangs from the left indent: at least as deep as it hangs.
                    if (left < -f) ind.setAttributeNS(W, "w:$leftAttr", (-f).toString())
                }
                // Taking a hanging indent away also takes back the left indent it needed.
                oldHanging != null && left == oldHanging -> ind.removeAttributeNS(W, leftAttr)
            }
        }
        op.pageBreakBefore?.let { on ->
            child(pPr, "pageBreakBefore")?.let { pPr.removeChild(it) }
            if (on) insertInOrder(pPr, document.createElementNS(W, "w:pageBreakBefore"), PPR_ORDER)
        }
        notes.add("paragraph p${op.para}: $op")
    }

    private fun runStyle(document: Document, p: Element, op: Op.RunStyle, notes: MutableList<String>) {
        if (op.end <= op.start) return
        splitAt(p, op.start)
        splitAt(p, op.end)
        var offset = 0
        var touched = 0
        for (seg in DocxReader.segments(p)) {
            val len = seg.length
            if (seg.isRun && len > 0 && offset >= op.start && offset + len <= op.end) {
                val rPr = child(seg.el, "rPr") ?: document.createElementNS(W, "w:rPr").also { seg.el.insertBefore(it, seg.el.firstChild) }
                op.font?.let { name ->
                    // A named font on every script slot; theme references would override it.
                    child(rPr, "rFonts")?.let { rPr.removeChild(it) }
                    val f = document.createElementNS(W, "w:rFonts")
                    for (slot in listOf("ascii", "hAnsi", "eastAsia", "cs")) f.setAttributeNS(W, "w:$slot", name)
                    insertInOrder(rPr, f, RPR_ORDER)
                }
                op.size?.let { size ->
                    for (tag in listOf("sz", "szCs")) {
                        child(rPr, tag)?.let { rPr.removeChild(it) }
                        val e = document.createElementNS(W, "w:$tag")
                        e.setAttributeNS(W, "w:val", size.toString())
                        insertInOrder(rPr, e, RPR_ORDER)
                    }
                }
                touched++
            }
            offset += len
        }
        notes.add("font=${op.font} size=${op.size} p${op.para} [${op.start},${op.end}): $touched run(s)")
    }

    /** Sets a run property explicitly on or off, so styles cannot override the user's choice. */
    private fun setRunProperty(document: Document, r: Element, prop: String, on: Boolean) {
        val (name, value) = when (prop) {
            "b" -> "b" to (if (on) null else "0")
            "i" -> "i" to (if (on) null else "0")
            "u" -> "u" to (if (on) "single" else "none")
            "h" -> "highlight" to (if (on) "yellow" else "none")
            else -> return
        }
        var rPr = child(r, "rPr")
        if (rPr == null) {
            rPr = document.createElementNS(W, "w:rPr")
            r.insertBefore(rPr, r.firstChild)
        }
        child(rPr!!, name)?.let { rPr.removeChild(it) }
        val el = document.createElementNS(W, "w:$name")
        if (value != null) el.setAttributeNS(W, "w:val", value)
        val rank = RPR_ORDER.indexOf(name)
        // Unknown children (rPrChange, extensions) sort last: new properties go before them.
        val before = elementChildren(rPr).firstOrNull { (RPR_ORDER.indexOf(it.localName).takeIf { i -> i >= 0 } ?: 1000) > rank }
        rPr.insertBefore(el, before)
    }

    /** Sets (or, for Normal, removes) the paragraph's style. pStyle is always pPr's first child. */
    private fun setStyle(document: Document, p: Element, styleId: String?) {
        var pPr = child(p, "pPr")
        pPr?.let { pp -> child(pp, "pStyle")?.let { pp.removeChild(it) } }
        if (styleId == null) return
        if (pPr == null) {
            pPr = document.createElementNS(W, "w:pPr")
            p.insertBefore(pPr, p.firstChild)
        }
        val el = document.createElementNS(W, "w:pStyle")
        el.setAttributeNS(W, "w:val", styleId)
        pPr!!.insertBefore(el, pPr.firstChild)
    }

    /**
     * Finds the document's own heading/title style ids by their built-in names ("heading 1"
     * — ids differ by Word language), adding a minimal definition when a document lacks one.
     */
    private class StyleIds(private val styles: Document?) {
        var added = false
            private set
        private val byName = HashMap<String, String>()
        private val ids = HashSet<String>()
        private val next = HashMap<String, String>()
        private val raw = HashMap<String, Element>()
        private var defaultId: String? = null

        init {
            styles?.documentElement?.let { root ->
                for (s in elementChildren(root).filter { it.localName == "style" && it.getAttributeNS(W, "type") == "paragraph" }) {
                    val id = s.getAttributeNS(W, "styleId")
                    ids.add(id)
                    raw[id] = s
                    child(s, "name")?.getAttributeNS(W, "val")?.lowercase()?.let { byName.putIfAbsent(it, id) }
                    child(s, "next")?.getAttributeNS(W, "val")?.takeIf { it.isNotEmpty() }?.let { next[id] = it }
                    if (s.getAttributeNS(W, "default").let { it == "1" || it == "true" }) defaultId = id
                }
            }
        }

        /** Whether paragraphs in style [id] are numbered by the style itself (w:numPr, through basedOn). */
        fun numbers(id: String): Boolean {
            val seen = HashSet<String>()
            var cur: String? = id
            while (cur != null && seen.add(cur)) {
                val s = raw[cur] ?: return false
                val numPr = child(s, "pPr")?.let { child(it, "numPr") }
                if (numPr != null) return (child(numPr, "numId")?.getAttributeNS(W, "val")?.toIntOrNull() ?: 0) > 0
                cur = child(s, "basedOn")?.getAttributeNS(W, "val")
            }
            return false
        }

        /** The style Word gives the paragraph after one in [id] (heading → Normal). */
        fun nextOf(id: String): String? = next[id]

        fun isDefault(id: String): Boolean = id == defaultId

        /** Null means "Normal": remove pStyle so the paragraph takes the default style. */
        fun idFor(kind: String): String? {
            val (name, fallbackId, outline, size) = when (kind) {
                "heading1" -> Quad("heading 1", "Heading1", 0, 32)
                "heading2" -> Quad("heading 2", "Heading2", 1, 28)
                "title" -> Quad("title", "Title", null, 56)
                else -> return null
            }
            byName[name]?.let { return it }
            val root = styles?.documentElement ?: return fallbackId
            var id = fallbackId
            var n = 2
            while (id in ids) id = "$fallbackId${n++}"
            val s = styles.createElementNS(W, "w:style")
            s.setAttributeNS(W, "w:type", "paragraph")
            s.setAttributeNS(W, "w:styleId", id)
            fun add(parent: Element, tag: String, value: String? = null): Element {
                val e = styles.createElementNS(W, "w:$tag")
                if (value != null) e.setAttributeNS(W, "w:val", value)
                parent.appendChild(e)
                return e
            }
            add(s, "name", name)
            byName["normal"]?.let { add(s, "basedOn", it); add(s, "next", it) }
            add(s, "qFormat")
            val pPr = add(s, "pPr")
            add(pPr, "keepNext")
            val spacing = add(pPr, "spacing")
            spacing.setAttributeNS(W, "w:before", "240")
            spacing.setAttributeNS(W, "w:after", "120")
            if (outline != null) add(pPr, "outlineLvl", outline.toString())
            val rPr = add(s, "rPr")
            add(rPr, "b")
            add(rPr, "sz", size.toString())
            root.appendChild(s)
            byName[name] = id
            ids.add(id)
            added = true
            return id
        }

        private data class Quad(val name: String, val id: String, val outline: Int?, val size: Int)
    }

    // ---------------------------------------------------------------- package helpers

    fun serialize(doc: Document): ByteArray {
        doc.xmlStandalone = true
        val t = TransformerFactory.newInstance().newTransformer()
        t.setOutputProperty(OutputKeys.ENCODING, "UTF-8")
        t.setOutputProperty(OutputKeys.INDENT, "no")
        t.setOutputProperty(OutputKeys.STANDALONE, "yes")
        val out = ByteArrayOutputStream()
        t.transform(DOMSource(doc), StreamResult(out))
        return out.toByteArray()
    }

    private fun putEntry(out: ZipOutputStream, original: ZipEntry, bytes: ByteArray) {
        val entry = ZipEntry(original.name)
        entry.time = original.time
        if (original.method == ZipEntry.STORED) {
            entry.method = ZipEntry.STORED
            entry.size = bytes.size.toLong()
            entry.compressedSize = bytes.size.toLong()
            entry.crc = CRC32().apply { update(bytes) }.value
        } else {
            entry.method = ZipEntry.DEFLATED
        }
        out.putNextEntry(entry)
        out.write(bytes)
        out.closeEntry()
    }

    private fun hashes(file: File): Map<String, String> {
        val out = HashMap<String, String>()
        ZipFile(file).use { zip ->
            for (entry in zip.entries()) {
                val md = MessageDigest.getInstance("SHA-256")
                md.update(DocxReader.readEntry(zip, entry))
                out[entry.name] = md.digest().joinToString("") { "%02x".format(it) }
            }
        }
        return out
    }

    /** Plain write over [dest]: creating and overwriting are allowed for plugins; deleting is not. */
    private fun copy(from: File, dest: File) {
        from.inputStream().use { input -> FileOutputStream(dest).use { input.copyTo(it) } }
    }

    /** `<stem>-edited.docx`, or `-edited-2`, `-3` … beside the original. */
    fun editedCopyName(src: File): File {
        val dir = src.parentFile ?: File(".")
        val stem = src.nameWithoutExtension
        var file = File(dir, "$stem-edited.docx")
        var n = 2
        while (file.exists()) file = File(dir, "$stem-edited-${n++}.docx")
        return file
    }
}
