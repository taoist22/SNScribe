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
    /** kind: "heading1" / "heading2" / "heading3" / "title"; null values leave that property alone. */
    data class StyleDef(
        val kind: String,
        val font: String? = null,
        val size: Int? = null,
        val bold: Boolean? = null,
        val italic: Boolean? = null,
        val align: String? = null,
        val line: Int? = null,
    )

    sealed class Op {
        /** prop: "b", "i", "u" or "h". */
        data class Format(val para: Int, val start: Int, val end: Int, val prop: String, val on: Boolean, val value: String? = null) : Op()

        /** kind: "heading1", "heading2", "heading3", "title", "quote" or "normal". */
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

        /**
         * The document's header or footer ([kind] "header" / "footer") on every page: one
         * paragraph aligned [align], holding [text] and, when [pageNumber], Word's PAGE field
         * after it. Empty text and no page number leaves it blank. Applies to every section.
         */
        data class HeaderFooter(val kind: String, val text: String, val pageNumber: Boolean, val align: String) : Op()

        /** Makes characters [start, end) of one paragraph a hyperlink to [url]. */
        data class Link(val para: Int, val start: Int, val end: Int, val url: String) : Op()

        /** Removes hyperlinks that overlap characters [start, end) (their text stays). */
        data class Unlink(val para: Int, val start: Int, val end: Int) : Op()

        /**
         * Accepts or rejects the tracked insertion or deletion [id] in one paragraph; [para]
         * −1 with [id] "*" does every text insertion and deletion in the document's paragraphs.
         */
        data class Revision(val para: Int, val id: String, val accept: Boolean) : Op()

        /**
         * A new comment [id] on the text from ([fromPara], [from]) to ([toPara], [to]), or a
         * reply to comment [parent] (its range is the parent's; the positions are ignored).
         */
        data class CommentAdd(
            val id: Int,
            val fromPara: Int,
            val from: Int,
            val toPara: Int,
            val to: Int,
            val text: String,
            val author: String,
            val initials: String,
            val date: String,
            val parent: Int? = null,
        ) : Op()

        /** Deletes these comments (a thread is its comment and all its replies). */
        data class CommentDelete(val ids: List<Int>) : Op()

        /**
         * Redefines paragraph styles (a paper format's headings): each [StyleDef] rewrites
         * that style's font, size, bold, italic, alignment, indent and line spacing, creating
         * the style if the document has none. Headings added later then come out right.
         */
        data class StyleDefs(val defs: List<StyleDef>) : Op()

        /** Cell [cell] of row [row] of table [table] (ordinal among the body's tables) gets the text [pieces] ("\n" = new paragraph). */
        data class TableCell(val table: Int, val row: Int, val cell: Int, val pieces: List<DocxReader.NotePiece>) : Op()

        /** A new empty row like row [row], above or [below] it. */
        data class TableRowAdd(val table: Int, val row: Int, val below: Boolean) : Op()

        /** Removes row [row] (never the last one). */
        data class TableRowDelete(val table: Int, val row: Int) : Op()

        /** A new table of [rows]×[cols] empty cells before paragraph [before]; the first row a [header] row. */
        data class TableInsert(val before: Int, val rows: Int, val cols: Int, val header: Boolean) : Op()

        /** A footnote [id] with the text [pieces], its number at character [at] of [para]. */
        data class FootnoteAdd(val para: Int, val at: Int, val id: Int, val pieces: List<DocxReader.NotePiece>) : Op()

        /** Footnote [id]'s text becomes [pieces] (its number stays). */
        data class FootnoteSet(val id: Int, val pieces: List<DocxReader.NotePiece>) : Op()

        /** Removes footnote [id] and its number in the text. */
        data class FootnoteDelete(val id: Int) : Op()

        /** A picture from the file [path] (PNG, JPEG, GIF or BMP) in the text at [at], [cx]×[cy] EMU; [alt] its description. */
        data class ImageAdd(val para: Int, val at: Int, val path: String, val cx: Long, val cy: Long, val alt: String = "") : Op()

        /** A handwritten note [id]: the PNG [png] ([width]×[height] px) in the right margin, anchored at [at]. */
        data class InkAdd(val para: Int, val at: Int, val id: String, val png: String, val width: Int, val height: Int) : Op()

        /** Removes handwritten note [id] (in paragraph [para]). */
        data class InkDelete(val para: Int, val id: String) : Op()
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
    /**
     * The package being edited: parts parsed on first use, parts changed in place, parts
     * added (each registered in [Content_Types].xml and, when it belongs to the document,
     * in document.xml.rels). Untouched parts are copied byte for byte.
     */
    class Pkg(private val zip: ZipFile?) {
        private val parsed = HashMap<String, Document>()
        val dirty = LinkedHashSet<String>()
        val added = LinkedHashMap<String, Document>()
        val addedBinary = LinkedHashMap<String, ByteArray>()
        val removed = LinkedHashSet<String>()

        fun part(name: String): Document? =
            added[name] ?: parsed[name] ?: zip?.getEntry(name)?.let { DocxReader.parse(DocxReader.readEntry(zip, it)) }?.also { parsed[name] = it }

        fun exists(name: String) = added.containsKey(name) || addedBinary.containsKey(name) ||
            (name !in removed && zip?.getEntry(name) != null)

        /** Adds a binary part (a picture); its extension gets a Default content type if it has none. */
        fun addBinary(name: String, bytes: ByteArray, extension: String, contentType: String) {
            addedBinary[name] = bytes
            val types = part(CONTENT_TYPES) ?: error("no $CONTENT_TYPES")
            val root = types.documentElement
            val has = elementChildren(root).any { it.localName == "Default" && it.getAttribute("Extension").equals(extension, ignoreCase = true) }
            if (!has) {
                val d = types.createElementNS(root.namespaceURI, "Default")
                d.setAttribute("Extension", extension)
                d.setAttribute("ContentType", contentType)
                root.insertBefore(d, root.firstChild)
                touch(CONTENT_TYPES)
            }
        }

        /** Removes a part from the package (one added this session simply goes). */
        fun remove(name: String) {
            if (added.remove(name) == null && addedBinary.remove(name) == null) removed.add(name)
        }

        /** Removes document relationship [id]. */
        fun unrelate(id: String) {
            val rels = part(DocxReader.DOCUMENT_RELS) ?: return
            elementChildren(rels.documentElement).filter { it.getAttribute("Id") == id }.forEach { rels.documentElement.removeChild(it) }
            touch(DocxReader.DOCUMENT_RELS)
        }

        /** Marks an existing part changed (it will be re-serialized). */
        fun touch(name: String) {
            if (name !in added) dirty.add(name)
        }

        /** Replaces an existing part's content wholesale. */
        fun replace(name: String, doc: Document) {
            if (name in added) added[name] = doc else parsed[name] = doc
            touch(name)
        }

        /** A free part name like word/header3.xml. */
        fun freeName(stem: String, ext: String = "xml"): String {
            // (Removed names are not reused: a part name maps to one picture per session.)
            var n = 1
            while (exists("$stem$n.$ext")) n++
            return "$stem$n.$ext"
        }

        /**
         * Adds a part, registers its content type, and relates it from the document
         * ([relType] null = not related). Returns the relationship id.
         */
        fun add(name: String, doc: Document, contentType: String, relType: String?): String? {
            added[name] = doc
            val types = part(CONTENT_TYPES) ?: error("no $CONTENT_TYPES")
            addOverride(types, "/$name", contentType)
            touch(CONTENT_TYPES)
            return relType?.let { relate(it, name.removePrefix("word/"), external = false) }
        }

        /** A new relationship from the document; returns its id. */
        fun relate(type: String, target: String, external: Boolean): String {
            val rels = part(DocxReader.DOCUMENT_RELS) ?: error("no ${DocxReader.DOCUMENT_RELS}")
            touch(DocxReader.DOCUMENT_RELS)
            return addRelationship(rels, type, target, external)
        }

        /** The target (as a part name) of the document relationship [id]. */
        fun target(id: String): String? {
            val rels = part(DocxReader.DOCUMENT_RELS) ?: return null
            var n = rels.documentElement.firstChild
            while (n != null) {
                if (n is Element && n.getAttribute("Id") == id) {
                    val t = n.getAttribute("Target")
                    return if (t.startsWith("/")) t.removePrefix("/") else "word/" + t.removePrefix("./")
                }
                n = n.nextSibling
            }
            return null
        }
    }

    fun save(src: File, ops: List<Op>, dest: File, workDir: File, expected: List<String> = emptyList()): Saved {
        require(src.length() <= DocxReader.MAX_DOCX_BYTES) { "too large: ${src.length()} bytes" }
        val effects = ArrayList<Op>()
        val notes = ArrayList<String>()
        val changed = ArrayList<String>()
        var addedNames: Set<String> = emptySet()
        var removedNames: Set<String> = emptySet()
        workDir.mkdirs()
        val temp = File(workDir, "saving-${System.nanoTime()}.docx")
        try {
            ZipFile(src).use { zip ->
                val pkg = Pkg(zip)
                val document = pkg.part(DOCUMENT_PART) ?: error("no $DOCUMENT_PART")
                val styles = pkg.part(STYLES_PART)
                val numberingPath = DocxReader.numberingPart(pkg.part(DocxReader.DOCUMENT_RELS))
                val numbering = pkg.part(numberingPath)
                val lists = Lists(numbering)
                val stylesChanged = apply(document, styles, ops, notes, lists, pkg, effects)
                if (ops.isNotEmpty()) pkg.touch(DOCUMENT_PART)
                if (stylesChanged && styles != null) pkg.touch(STYLES_PART)
                if (lists.changed) {
                    if (numbering != null) {
                        pkg.touch(numberingPath)
                    } else {
                        // A document with no lists yet: add the part and register it.
                        check(!pkg.exists(numberingPath)) { "unreadable $numberingPath" }
                        pkg.add(numberingPath, lists.doc, NUMBERING_TYPE, DocxReader.REL_NUMBERING)
                    }
                }

                ZipOutputStream(FileOutputStream(temp)).use { out ->
                    for (entry in zip.entries()) {
                        DocxReader.checkEntryName(entry.name)
                        if (entry.name in pkg.removed) continue
                        val bytes = if (entry.name in pkg.dirty) serialize(pkg.part(entry.name)!!) else DocxReader.readEntry(zip, entry)
                        putEntry(out, entry, bytes)
                    }
                    for ((name, doc) in pkg.added) putEntry(out, ZipEntry(name).apply { method = ZipEntry.DEFLATED }, serialize(doc))
                    for ((name, bytes) in pkg.addedBinary) putEntry(out, ZipEntry(name).apply { method = ZipEntry.DEFLATED }, bytes)
                }
                changed.addAll(pkg.dirty)
                changed.addAll(pkg.added.keys)
                changed.addAll(pkg.addedBinary.keys)
                addedNames = pkg.added.keys + pkg.addedBinary.keys
                removedNames = pkg.removed.toSet()
            }
            verify(src, temp, changed, effects, expected, addedNames, removedNames)
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
    private fun verify(src: File, written: File, changed: List<String>, ops: List<Op>, expected: List<String>, added: Set<String>, removed: Set<String> = emptySet()) {
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
        check(hb.keys == ha.keys - removed + added) { "verify: package parts differ: ${ha.keys - hb.keys} / ${hb.keys - ha.keys}" }
        for ((name, hash) in ha) {
            if (name !in changed && name !in removed) check(hb[name] == hash) { "verify: $name changed but was not edited" }
        }
    }

    // ---------------------------------------------------------------- apply

    /** Applies [ops] to the DOM. Returns whether styles.xml was changed (a style was added). */
    fun apply(
        document: Document,
        styles: Document?,
        ops: List<Op>,
        notes: MutableList<String>,
        lists: Lists = Lists(null),
        pkg: Pkg = Pkg(null),
        effects: MutableList<Op>? = null,
    ): Boolean {
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
            if (op is Op.HeaderFooter) {
                headerFooter(document, pkg, op, notes)
                continue
            }
            if (op is Op.StyleDefs) {
                if (styles != null) {
                    for (d in op.defs) styleIds.redefine(d)
                    stylesTouched = true
                    notes.add("styles redefined: ${op.defs.map { it.kind }}")
                }
                continue
            }
            if (op is Op.CommentAdd) {
                addComment(document, paragraphs, pkg, op, notes)
                continue
            }
            if (op is Op.CommentDelete) {
                deleteComments(document, pkg, op, notes)
                continue
            }
            if (op is Op.TableCell || op is Op.TableRowAdd || op is Op.TableRowDelete || op is Op.TableInsert) {
                // Tables hold no body paragraphs: paragraph addresses and texts are unchanged.
                table(document, paragraphs, op, notes)
                continue
            }
            if (op is Op.FootnoteAdd || op is Op.FootnoteSet || op is Op.FootnoteDelete) {
                // Body text changes (a number added or removed), as plain splices, for verify.
                val before = paragraphs.map { paragraphText(it) }
                when (op) {
                    is Op.FootnoteAdd -> {
                        val p = checkNotNull(paragraphs.getOrNull(op.para)) { "no paragraph ${op.para} for $op" }
                        if (styles != null && footnoteStyles(styles)) stylesTouched = true
                        addFootnote(document, p, pkg, styles, op, notes)
                    }
                    is Op.FootnoteSet -> setFootnote(document, pkg, styles, op.id, op.pieces, notes)
                    is Op.FootnoteDelete -> deleteFootnote(document, pkg, op, notes)
                    else -> {}
                }
                paragraphs.forEachIndexed { i, p ->
                    val now = paragraphText(p)
                    if (now != before[i]) effects?.add(Op.Text(i, 0, before[i].length, now))
                }
                continue
            }
            if (op is Op.ImageAdd) {
                val p = checkNotNull(paragraphs.getOrNull(op.para)) { "no paragraph ${op.para} for $op" }
                val old = paragraphText(p)
                addImage(document, p, pkg, op, notes)
                effects?.add(Op.Text(op.para, 0, old.length, paragraphText(p)))
                continue
            }
            if (op is Op.InkAdd || op is Op.InkDelete) {
                val i = if (op is Op.InkAdd) op.para else (op as Op.InkDelete).para
                val p = checkNotNull(paragraphs.getOrNull(i)) { "no paragraph $i for $op" }
                val old = paragraphText(p)
                if (op is Op.InkAdd) addInk(document, p, pkg, op, notes) else deleteInk(document, pkg, op as Op.InkDelete, notes)
                val new = paragraphText(p)
                if (old != new) effects?.add(Op.Text(i, 0, old.length, new))
                continue
            }
            if (op is Op.Revision) {
                // Text changes, as plain splices, for verify.
                val targets = if (op.para < 0) paragraphs.indices.toList() else listOf(op.para)
                for (i in targets) {
                    val p = checkNotNull(paragraphs.getOrNull(i)) { "no paragraph $i for $op" }
                    val old = paragraphText(p)
                    val done = revision(p, if (op.para < 0) null else op.id, op.accept)
                    check(done > 0 || op.para < 0) { "no tracked change ${op.id} in paragraph ${op.para}" }
                    val new = paragraphText(p)
                    if (old != new) effects?.add(Op.Text(i, 0, old.length, new))
                }
                notes.add("revision ${if (op.accept) "accept" else "reject"} ${op.id} p${op.para}")
                continue
            }
            effects?.add(op)
            val index = when (op) {
                is Op.Format -> op.para
                is Op.Style -> op.para
                is Op.Text -> op.para
                is Op.Split -> op.para
                is Op.Join -> op.para
                is Op.ListItem -> op.para
                is Op.RunStyle -> op.para
                is Op.ParaProps -> op.para
                is Op.Link -> op.para
                is Op.Unlink -> op.para
                is Op.PageSetup, is Op.Defaults, is Op.HeaderFooter, is Op.Revision, is Op.CommentAdd, is Op.CommentDelete, is Op.InkAdd, is Op.InkDelete, is Op.ImageAdd, is Op.StyleDefs,
                is Op.FootnoteAdd, is Op.FootnoteSet, is Op.FootnoteDelete,
                is Op.TableCell, is Op.TableRowAdd, is Op.TableRowDelete, is Op.TableInsert -> error("unreachable")
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
                is Op.Link -> link(document, p, op, pkg, styleIds, notes)
                is Op.Unlink -> unlink(p, op, notes)
                is Op.PageSetup, is Op.Defaults, is Op.HeaderFooter, is Op.Revision, is Op.CommentAdd, is Op.CommentDelete, is Op.InkAdd, is Op.InkDelete, is Op.ImageAdd, is Op.StyleDefs,
                is Op.FootnoteAdd, is Op.FootnoteSet, is Op.FootnoteDelete,
                is Op.TableCell, is Op.TableRowAdd, is Op.TableRowDelete, is Op.TableInsert -> {}
            }
            // Splits and joins renumber the paragraphs after them.
            if (op is Op.Split || op is Op.Join) paragraphs = bodyParagraphs(document)
        }
        return styleIds.added || stylesTouched
    }

    private fun paragraphText(p: Element): String = DocxReader.segments(p).joinToString("") { if (it.isRun) DocxReader.runText(it.el) else DocxReader.OBJECT.toString() }

    /**
     * Accepts or rejects tracked insertions and deletions in [p]'s content ([id] null = all
     * of them). Accepting an insertion or rejecting a deletion keeps its text as ordinary
     * text; the other two remove it. Returns how many changes were resolved.
     */
    private fun revision(p: Element, id: String?, accept: Boolean): Int {
        val found = ArrayList<Element>()
        fun collect(parent: Element) {
            for (el in elementChildren(parent)) {
                if (el.namespaceURI != W || el.localName == "pPr") continue
                if (el.localName in REVISION_TAGS && (id == null || el.getAttributeNS(W, "id") == id)) found.add(el)
                collect(el)
            }
        }
        collect(p)
        // Innermost first: a deletion inside an insertion is resolved before the insertion.
        for (el in found.asReversed()) {
            if (el.parentNode == null) continue
            val inserted = el.localName == "ins" || el.localName == "moveTo"
            if (inserted == accept) {
                // Keep the text: unwrap, and deleted text becomes text again.
                if (!inserted) renameDeleted(el)
                while (el.firstChild != null) el.parentNode.insertBefore(el.firstChild, el)
            }
            el.parentNode.removeChild(el)
        }
        return found.size
    }

    private val REVISION_TAGS = setOf("ins", "del", "moveTo", "moveFrom")

    private fun renameDeleted(el: Element) {
        val doc = el.ownerDocument
        for ((from, to) in listOf("delText" to "w:t", "delInstrText" to "w:instrText")) {
            val list = el.getElementsByTagNameNS(W, from)
            val items = (0 until list.length).map { list.item(it) }
            for (n in items) doc.renameNode(n, W, to)
        }
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
                setRunProperty(document, seg.el, op.prop, op.on, op.value)
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
        // Only the first half starts a new page: the rest continues on it (CT: pressing Enter
        // in a paragraph that started a page added another page break). A style's page break
        // stays with the style.
        child(p, "pPr")?.let { pPr -> child(pPr, "pageBreakBefore")?.let { pPr.removeChild(it) } }

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
            // A list item takes its indentation from the list: the paragraph's own (a copied
            // first-line indent, say) would override the list's and push the text away.
            if (numId > 0) child(pPr!!, "ind")?.let { pPr!!.removeChild(it) }
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

    private fun addRelationship(rels: Document, type: String, target: String, external: Boolean = false): String {
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
        if (external) r.setAttribute("TargetMode", "External")
        root.appendChild(r)
        return "rId$i"
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

    // ---------------------------------------------------------------- handwritten margin notes

    private const val REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    private const val EMU_PER_INCH = 914400L

    /**
     * A handwritten note in the right margin: the picture [Op.InkAdd.png] as a floating
     * image anchored at character [Op.InkAdd.at] — beside that line, moving with it, not
     * pushing text (wrapNone). Word shows it (Office 365 checked, 2026-09-26); its anchor is
     * one U+FFFC in the paragraph's text, like any drawing.
     */
    private fun addInk(document: Document, p: Element, pkg: Pkg, op: Op.InkAdd, notes: MutableList<String>) {
        val bytes = File(op.png).readBytes()
        check(bytes.size > 8 && bytes[1] == 'P'.code.toByte()) { "not a PNG: ${op.png}" }
        val part = pkg.freeName("word/media/sndocx-ink-", "png")
        pkg.addBinary(part, bytes, "png", "image/png")
        val rId = pkg.relate(REL_IMAGE, part.removePrefix("word/"), external = false)

        // As wide as the right margin allows (at least 0.8″, at most 1.6″), height to scale.
        val right = child(child(document.documentElement, "body") ?: document.documentElement, "sectPr")
            ?.let { child(it, "pgMar") }?.getAttributeNS(W, "right")?.toIntOrNull() ?: 1440
        val inches = (right / 1440.0 - 0.1).coerceIn(0.8, 1.6)
        val cx = (inches * EMU_PER_INCH).toLong()
        val cy = (cx * op.height / maxOf(1, op.width)).coerceAtLeast(EMU_PER_INCH / 10)
        // Notes never overlap: Word moves floating pictures apart when allowOverlap is off (set
        // on every DOCX note, older ones too), and notes beside one paragraph are stacked here
        // as well, for apps that ignore that setting.
        var below = 0L
        val drawingsNow = document.getElementsByTagNameNS(W, "drawing")
        for (i in 0 until drawingsNow.length) {
            val d = drawingsNow.item(i) as Element
            if (DocxReader.inkNoteId(d) == null) continue
            val anchor = d.getElementsByTagNameNS(DocxReader.WP, "anchor").item(0) as? Element ?: continue
            anchor.setAttribute("allowOverlap", "0")
            var n: org.w3c.dom.Node? = d
            while (n != null && n !== p) n = n.parentNode
            if (n === p) {
                val ext = anchor.getElementsByTagNameNS(DocxReader.WP, "extent").item(0) as? Element
                val off = (anchor.getElementsByTagNameNS(DocxReader.WP, "posOffset").item(1) as? Element)?.textContent?.trim()?.toLongOrNull() ?: 0L
                below = maxOf(below, off + (ext?.getAttribute("cy")?.toLongOrNull() ?: 0L) + EMU_PER_INCH / 20)
            }
        }
        var maxDocPr = 0
        val prs = document.getElementsByTagNameNS(DocxReader.WP, "docPr")
        for (i in 0 until prs.length) maxDocPr = maxOf(maxDocPr, (prs.item(i) as Element).getAttribute("id").toIntOrNull() ?: 0)
        val docPr = maxDocPr + 1
        val xml = """<w:r xmlns:w="$W" xmlns:wp="${DocxReader.WP}" xmlns:a="${DocxReader.A}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="$R_NS"><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251659264 + docPr}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="rightMargin"><wp:posOffset>45720</wp:posOffset></wp:positionH><wp:positionV relativeFrom="line"><wp:posOffset>$below</wp:posOffset></wp:positionV><wp:extent cx="$cx" cy="$cy"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="$docPr" name="${DocxReader.INK_NAME}${op.id}" descr="Handwritten note"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="$docPr" name="Handwritten note ${op.id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$cx" cy="$cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"""
        val run = document.importNode(DocxReader.parse(xml.toByteArray()).documentElement, true) as Element
        // Namespaces are declared on the document root already, or the serializer adds them.
        for (prefix in listOf("w", "wp", "a", "pic", "r")) run.removeAttributeNS("http://www.w3.org/2000/xmlns/", prefix)
        placeAt(p, op.at, run, before = true)
        notes.add("ink note ${op.id} at p${op.para}:${op.at} → $part (${"%.2f".format(inches)}″)")
    }

    // ---------------------------------------------------------------- tables

    private fun bodyTables(document: Document): List<Element> {
        val body = child(document.documentElement, "body") ?: return emptyList()
        return elementChildren(body).filter { it.localName == "tbl" && it.namespaceURI == W }
    }

    private fun table(document: Document, paragraphs: List<Element>, op: Op, notes: MutableList<String>) {
        if (op is Op.TableInsert) {
            insertTable(document, paragraphs, op, notes)
            return
        }
        val index = when (op) {
            is Op.TableCell -> op.table
            is Op.TableRowAdd -> op.table
            is Op.TableRowDelete -> op.table
            else -> error("not a table edit: $op")
        }
        val tbl = checkNotNull(bodyTables(document).getOrNull(index)) { "no table $index" }
        val rows = elementChildren(tbl).filter { it.localName == "tr" }
        when (op) {
            is Op.TableCell -> {
                val tr = checkNotNull(rows.getOrNull(op.row)) { "table $index has no row ${op.row}" }
                val tc = checkNotNull(elementChildren(tr).filter { it.localName == "tc" }.getOrNull(op.cell)) { "row ${op.row} has no cell ${op.cell}" }
                check(elementChildren(tc).none { it.localName == "tbl" }) { "that cell holds a table of its own" }
                setCell(document, tc, op.pieces)
                notes.add("table $index cell ${op.row},${op.cell}")
            }
            is Op.TableRowAdd -> {
                val tr = checkNotNull(rows.getOrNull(op.row)) { "table $index has no row ${op.row}" }
                val copy = tr.cloneNode(true) as Element
                // A new row is an ordinary row: not a repeated header, merges from above undone.
                child(copy, "trPr")?.let { trPr -> child(trPr, "tblHeader")?.let { trPr.removeChild(it) } }
                for (tc in elementChildren(copy).filter { it.localName == "tc" }) {
                    child(tc, "tcPr")?.let { tcPr -> child(tcPr, "vMerge")?.let { tcPr.removeChild(it) } }
                    elementChildren(tc).filter { it.localName == "tbl" }.forEach { tc.removeChild(it) }
                    setCell(document, tc, emptyList())
                }
                tbl.insertBefore(copy, if (op.below) tr.nextSibling else tr)
                notes.add("table $index row added ${if (op.below) "below" else "above"} ${op.row}")
            }
            is Op.TableRowDelete -> {
                check(rows.size > 1) { "a table keeps at least one row" }
                val tr = checkNotNull(rows.getOrNull(op.row)) { "table $index has no row ${op.row}" }
                tbl.removeChild(tr)
                notes.add("table $index row ${op.row} deleted")
            }
            else -> {}
        }
    }

    /**
     * A cell's text: its paragraphs replaced by one per line of [pieces], keeping the first
     * paragraph's settings and the first run's look (font, size) — bold and italic as given.
     */
    private fun setCell(document: Document, tc: Element, pieces: List<DocxReader.NotePiece>) {
        val paras = elementChildren(tc).filter { it.localName == "p" }
        val first = paras.firstOrNull()
        val pPr = first?.let { child(it, "pPr") }?.cloneNode(true) as Element?
        val base = first?.let { p -> elementChildren(p).firstOrNull { it.localName == "r" }?.let { child(it, "rPr") } }?.cloneNode(true) as Element?
        base?.let { r -> elementChildren(r).filter { it.localName in setOf("b", "bCs", "i", "iCs", "rPrChange") }.forEach { r.removeChild(it) } }
        paras.forEach { tc.removeChild(it) }
        val lines = ArrayList<MutableList<DocxReader.NotePiece>>().apply { add(ArrayList()) }
        for (piece in pieces) {
            val parts = piece.text.split('\n')
            parts.forEachIndexed { i, part ->
                if (i > 0) lines.add(ArrayList())
                if (part.isNotEmpty()) lines.last().add(piece.copy(text = part))
            }
        }
        for (line in lines) {
            val p = document.createElementNS(W, "w:p")
            pPr?.let { p.appendChild(it.cloneNode(true)) }
            for (piece in line) {
                val r = document.createElementNS(W, "w:r")
                val rPr = (base?.cloneNode(true) as Element?) ?: document.createElementNS(W, "w:rPr")
                if (piece.bold) insertInOrder(rPr, document.createElementNS(W, "w:b"), RPR_ORDER)
                if (piece.italic) insertInOrder(rPr, document.createElementNS(W, "w:i"), RPR_ORDER)
                if (rPr.hasChildNodes()) r.appendChild(rPr)
                val t = document.createElementNS(W, "w:t")
                t.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
                t.textContent = piece.text
                r.appendChild(t)
                p.appendChild(r)
            }
            tc.appendChild(p)
        }
    }

    /**
     * A plain table before body paragraph [Op.TableInsert.before], as wide as the text, equal
     * columns — ruled as APA tables are: a line above and below the table and under the
     * header row, no vertical lines.
     */
    private fun insertTable(document: Document, paragraphs: List<Element>, op: Op.TableInsert, notes: MutableList<String>) {
        val at = checkNotNull(paragraphs.getOrNull(op.before)) { "no paragraph ${op.before} to put a table before" }
        check(op.rows in 1..50 && op.cols in 1..12) { "table size out of range" }
        val body = checkNotNull(child(document.documentElement, "body"))
        val sect = child(body, "sectPr")
        val pgSz = sect?.let { child(it, "pgSz") }
        val pgMar = sect?.let { child(it, "pgMar") }
        fun twips(e: Element?, a: String, dflt: Int) = e?.getAttributeNS(W, a)?.toIntOrNull() ?: dflt
        val textW = (twips(pgSz, "w", 12240) - twips(pgMar, "left", 1440) - twips(pgMar, "right", 1440)).coerceAtLeast(1440)
        val colW = textW / op.cols
        val line = "w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\""
        val grid = (1..op.cols).joinToString("") { "<w:gridCol w:w=\"$colW\"/>" }
        fun row(header: Boolean) = buildString {
            append("<w:tr>")
            if (header) append("<w:trPr><w:tblHeader/></w:trPr>")
            repeat(op.cols) {
                append("<w:tc><w:tcPr><w:tcW w:w=\"$colW\" w:type=\"dxa\"/>")
                if (header) append("<w:tcBorders><w:bottom $line/></w:tcBorders>")
                append("</w:tcPr><w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/><w:ind w:left=\"0\" w:firstLine=\"0\"/></w:pPr></w:p></w:tc>")
            }
            append("</w:tr>")
        }
        val xml = "<w:tbl xmlns:w=\"$W\"><w:tblPr><w:tblW w:w=\"$textW\" w:type=\"dxa\"/><w:tblBorders><w:top $line/><w:bottom $line/></w:tblBorders>" +
            "<w:tblLayout w:type=\"fixed\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr>" +
            "<w:tblGrid>$grid</w:tblGrid>" + (0 until op.rows).joinToString("") { row(op.header && it == 0) } + "</w:tbl>"
        val tbl = document.importNode(DocxReader.parse(xml.toByteArray()).documentElement, true) as Element
        tbl.removeAttributeNS("http://www.w3.org/2000/xmlns/", "w")
        at.parentNode.insertBefore(tbl, at)
        notes.add("table ${op.rows}×${op.cols} before p${op.before}")
    }

    // ---------------------------------------------------------------- footnotes

    private const val FOOTNOTES_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

    /** The footnotes part, created (with Word's two separators) when [create] and missing. */
    private fun footnotesPath(pkg: Pkg, create: Boolean): String? {
        DocxReader.relTarget(pkg.part(DocxReader.DOCUMENT_RELS), DocxReader.REL_FOOTNOTES)?.takeIf { pkg.exists(it) }?.let { return it }
        if (!create) return null
        val name = "word/footnotes.xml"
        check(!pkg.exists(name)) { "unreadable $name" }
        val xml = """<w:footnotes xmlns:w="$W"><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>"""
        pkg.add(name, DocxReader.parse(xml.toByteArray()), FOOTNOTES_TYPE, DocxReader.REL_FOOTNOTES)
        return name
    }

    /** A style's id by its built-in name ("footnote text"), or null. */
    private fun styleIdByName(styles: Document, type: String, name: String): String? =
        elementChildren(styles.documentElement).firstOrNull {
            it.localName == "style" && it.getAttributeNS(W, "type") == type && child(it, "name")?.getAttributeNS(W, "val")?.equals(name, ignoreCase = true) == true
        }?.getAttributeNS(W, "styleId")

    /**
     * Word's footnote styles, added to styles.xml when missing: "footnote reference" (a raised
     * number) and "footnote text" (10 pt, single spaced). Returns whether styles.xml changed.
     */
    private fun footnoteStyles(styles: Document): Boolean {
        var changed = false
        val root = styles.documentElement
        if (styleIdByName(styles, "character", "footnote reference") == null) {
            root.appendChild(styles.importNode(DocxReader.parse("""<w:style xmlns:w="$W" w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>""".toByteArray()).documentElement, true))
            changed = true
        }
        if (styleIdByName(styles, "paragraph", "footnote text") == null) {
            val base = elementChildren(root).firstOrNull { it.localName == "style" && it.getAttributeNS(W, "type") == "paragraph" && it.getAttributeNS(W, "default").let { d -> d == "1" || d == "true" } }
                ?.getAttributeNS(W, "styleId")
            val basedOn = base?.let { "<w:basedOn w:val=\"$it\"/>" } ?: ""
            root.appendChild(styles.importNode(DocxReader.parse("""<w:style xmlns:w="$W" w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>$basedOn<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>""".toByteArray()).documentElement, true))
            changed = true
        }
        for (s in elementChildren(root).filter { it.localName == "style" }) s.removeAttributeNS("http://www.w3.org/2000/xmlns/", "w")
        return changed
    }

    /** Runs for a footnote's text: italics and bold as given; "\n" becomes a space (one paragraph). */
    private fun noteRuns(doc: Document, pieces: List<DocxReader.NotePiece>): List<Element> = pieces.filter { it.text.isNotEmpty() }.map { piece ->
        doc.createElementNS(W, "w:r").also { r ->
            if (piece.italic || piece.bold) {
                val rPr = doc.createElementNS(W, "w:rPr")
                if (piece.bold) rPr.appendChild(doc.createElementNS(W, "w:b"))
                if (piece.italic) rPr.appendChild(doc.createElementNS(W, "w:i"))
                r.appendChild(rPr)
            }
            val t = doc.createElementNS(W, "w:t")
            t.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
            t.textContent = piece.text.replace('\n', ' ')
            r.appendChild(t)
        }
    }

    /** The number run, as Word writes it: the footnote reference style, else raised text. */
    private fun numberRun(doc: Document, styleId: String?, mark: String, id: Int?): Element {
        val r = doc.createElementNS(W, "w:r")
        val rPr = doc.createElementNS(W, "w:rPr")
        if (styleId != null) {
            rPr.appendChild(doc.createElementNS(W, "w:rStyle").also { it.setAttributeNS(W, "w:val", styleId) })
        } else {
            rPr.appendChild(doc.createElementNS(W, "w:vertAlign").also { it.setAttributeNS(W, "w:val", "superscript") })
        }
        r.appendChild(rPr)
        r.appendChild(doc.createElementNS(W, "w:$mark").also { e -> id?.let { e.setAttributeNS(W, "w:id", it.toString()) } })
        return r
    }

    /** Footnote [Op.FootnoteAdd.id]: its number at the given place, its text in the footnotes part. */
    private fun addFootnote(document: Document, p: Element, pkg: Pkg, styles: Document?, op: Op.FootnoteAdd, notes: MutableList<String>) {
        val path = checkNotNull(footnotesPath(pkg, create = true))
        val part = checkNotNull(pkg.part(path)) { "unreadable $path" }
        val all = elementChildren(part.documentElement).filter { it.localName == "footnote" }
        check(all.none { it.getAttributeNS(W, "id") == op.id.toString() }) { "footnote ${op.id} exists already" }
        val refStyle = styles?.let { styleIdByName(it, "character", "footnote reference") }
        placeAt(p, op.at, numberRun(document, refStyle, "footnoteReference", op.id), before = true)
        val note = part.createElementNS(W, "w:footnote")
        note.setAttributeNS(W, "w:id", op.id.toString())
        part.documentElement.appendChild(note)
        fillFootnote(part, note, styles, op.pieces)
        pkg.touch(path)
        notes.add("footnote ${op.id} at p${op.para}:${op.at}")
    }

    /** [note]'s content becomes one paragraph: its number, a space, then [pieces]. */
    private fun fillFootnote(part: Document, note: Element, styles: Document?, pieces: List<DocxReader.NotePiece>) {
        val first = elementChildren(note).firstOrNull { it.localName == "p" }
        val keptPPr = first?.let { child(it, "pPr") }?.cloneNode(true) as Element?
        while (note.firstChild != null) note.removeChild(note.firstChild)
        val p = part.createElementNS(W, "w:p")
        val pPr = keptPPr?.let { part.importNode(it, true) as Element } ?: part.createElementNS(W, "w:pPr").also { pp ->
            styles?.let { styleIdByName(it, "paragraph", "footnote text") }?.let { id ->
                pp.appendChild(part.createElementNS(W, "w:pStyle").also { it.setAttributeNS(W, "w:val", id) })
            }
        }
        if (pPr.hasChildNodes()) p.appendChild(pPr)
        p.appendChild(numberRun(part, styles?.let { styleIdByName(it, "character", "footnote reference") }, "footnoteRef", null))
        p.appendChild(noteRuns(part, listOf(DocxReader.NotePiece(" "))).single())
        for (r in noteRuns(part, pieces)) p.appendChild(r)
        note.appendChild(p)
    }

    private fun setFootnote(document: Document, pkg: Pkg, styles: Document?, id: Int, pieces: List<DocxReader.NotePiece>, notes: MutableList<String>) {
        val path = checkNotNull(footnotesPath(pkg, create = false)) { "no footnotes in this document" }
        val part = checkNotNull(pkg.part(path))
        val note = elementChildren(part.documentElement).firstOrNull { it.localName == "footnote" && it.getAttributeNS(W, "id") == id.toString() }
            ?: error("no footnote $id")
        fillFootnote(part, note, styles, pieces)
        pkg.touch(path)
        notes.add("footnote $id text")
    }

    private fun deleteFootnote(document: Document, pkg: Pkg, op: Op.FootnoteDelete, notes: MutableList<String>) {
        val refs = document.getElementsByTagNameNS(W, "footnoteReference")
        val mine = (0 until refs.length).map { refs.item(it) as Element }.filter { it.getAttributeNS(W, "id") == op.id.toString() }
        check(mine.isNotEmpty()) { "no footnote ${op.id} in the text" }
        for (ref in mine) {
            val run = ref.parentNode as Element
            run.removeChild(ref)
            if (elementChildren(run).all { it.localName == "rPr" }) run.parentNode.removeChild(run)
        }
        footnotesPath(pkg, create = false)?.let { path ->
            val part = checkNotNull(pkg.part(path))
            elementChildren(part.documentElement).filter { it.localName == "footnote" && it.getAttributeNS(W, "id") == op.id.toString() }
                .forEach { part.documentElement.removeChild(it) }
            pkg.touch(path)
        }
        notes.add("footnote ${op.id} deleted")
    }

    /** Picture formats a document may carry, by file extension: the content type Word expects. */
    private val PICTURE_TYPES = mapOf("png" to "image/png", "jpg" to "image/jpeg", "jpeg" to "image/jpeg", "gif" to "image/gif", "bmp" to "image/bmp")

    /**
     * A picture in the text (inline, as Word's Insert ▸ Pictures puts it): the file [Op.ImageAdd.path]
     * copied into word/media, at character [Op.ImageAdd.at], [Op.ImageAdd.cx]×[Op.ImageAdd.cy] EMU.
     * It is one U+FFFC in the paragraph's text, like any picture.
     */
    private fun addImage(document: Document, p: Element, pkg: Pkg, op: Op.ImageAdd, notes: MutableList<String>) {
        val ext = op.path.substringAfterLast('.', "").lowercase()
        val type = checkNotNull(PICTURE_TYPES[ext]) { "not a picture Word can hold: ${op.path}" }
        val bytes = File(op.path).readBytes()
        check(bytes.size in 1..25_000_000) { "picture too large (${bytes.size} bytes)" }
        check(op.cx in 1..(40 * EMU_PER_INCH) && op.cy in 1..(40 * EMU_PER_INCH)) { "picture size out of range" }
        val part = pkg.freeName("word/media/sndocx-picture-", ext)
        pkg.addBinary(part, bytes, ext, type)
        val rId = pkg.relate(REL_IMAGE, part.removePrefix("word/"), external = false)
        var maxDocPr = 0
        val prs = document.getElementsByTagNameNS(DocxReader.WP, "docPr")
        for (i in 0 until prs.length) maxDocPr = maxOf(maxDocPr, (prs.item(i) as Element).getAttribute("id").toIntOrNull() ?: 0)
        val docPr = maxDocPr + 1
        val alt = op.alt.replace("&", "&amp;").replace("<", "&lt;").replace("\"", "&quot;")
        val xml = """<w:r xmlns:w="$W" xmlns:wp="${DocxReader.WP}" xmlns:a="${DocxReader.A}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="$R_NS"><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${op.cx}" cy="${op.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="$docPr" name="Picture $docPr" descr="$alt"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="$docPr" name="Picture $docPr"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${op.cx}" cy="${op.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"""
        val run = document.importNode(DocxReader.parse(xml.toByteArray()).documentElement, true) as Element
        for (prefix in listOf("w", "wp", "a", "pic", "r")) run.removeAttributeNS("http://www.w3.org/2000/xmlns/", prefix)
        placeAt(p, op.at, run, before = true)
        notes.add("picture at p${op.para}:${op.at} → $part (${op.cx}×${op.cy} EMU)")
    }

    /** Removes handwritten note [Op.InkDelete.id]: its anchor, and its picture when nothing else shows it. */
    private fun deleteInk(document: Document, pkg: Pkg, op: Op.InkDelete, notes: MutableList<String>) {
        val drawings = document.getElementsByTagNameNS(W, "drawing")
        val d = (0 until drawings.length).map { drawings.item(it) as Element }.firstOrNull { DocxReader.inkNoteId(it) == op.id }
            ?: error("no handwritten note ${op.id}")
        val rId = (d.getElementsByTagNameNS(DocxReader.A, "blip").item(0) as? Element)?.getAttributeNS(R_NS, "embed")
        val run = d.parentNode as Element
        run.removeChild(d)
        if (elementChildren(run).all { it.localName == "rPr" }) run.parentNode.removeChild(run)
        if (rId != null) {
            val blips = document.getElementsByTagNameNS(DocxReader.A, "blip")
            val stillUsed = (0 until blips.length).any { (blips.item(it) as Element).getAttributeNS(R_NS, "embed") == rId }
            if (!stillUsed) pkg.target(rId)?.let { part -> pkg.unrelate(rId); pkg.remove(part) }
        }
        notes.add("ink note ${op.id} deleted")
    }

    // ---------------------------------------------------------------- comments

    private const val COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    private const val COMMENTS_EXT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    private const val REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
    private const val REL_COMMENTS_EXTENSIBLE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
    private const val W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
    private const val W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
    private const val XMLNS = "http://www.w3.org/2000/xmlns/"
    private const val MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

    private fun newPart(qualifiedRoot: String, ns: String, extra: List<Pair<String, String>>, ignorable: String): Document {
        val doc = javax.xml.parsers.DocumentBuilderFactory.newInstance().apply { isNamespaceAware = true }.newDocumentBuilder().newDocument()
        val root = doc.createElementNS(ns, qualifiedRoot)
        for ((prefix, uri) in extra) root.setAttributeNS(XMLNS, "xmlns:$prefix", uri)
        root.setAttributeNS(XMLNS, "xmlns:mc", MC)
        root.setAttributeNS(MC, "mc:Ignorable", ignorable)
        doc.appendChild(root)
        return doc
    }

    /** Declares prefix→uri on the root when missing, and lists it as ignorable (Word 2010+ markup). */
    private fun declare(root: Element, prefix: String, uri: String) {
        if (root.getAttributeNodeNS(XMLNS, prefix) == null) root.setAttributeNS(XMLNS, "xmlns:$prefix", uri)
        if (root.getAttributeNodeNS(XMLNS, "mc") == null) root.setAttributeNS(XMLNS, "xmlns:mc", MC)
        val ignorable = root.getAttributeNS(MC, "Ignorable").split(' ').filter { it.isNotEmpty() }
        if (prefix !in ignorable) root.setAttributeNS(MC, "mc:Ignorable", (ignorable + prefix).joinToString(" "))
    }

    private fun commentsPath(pkg: Pkg, create: Boolean): String? {
        DocxReader.relTarget(pkg.part(DocxReader.DOCUMENT_RELS), DocxReader.REL_COMMENTS)?.takeIf { pkg.exists(it) }?.let { return it }
        if (!create) return null
        val name = "word/comments.xml"
        check(!pkg.exists(name)) { "unreadable $name" }
        pkg.add(
            name,
            newPart("w:comments", W, listOf("w" to W, "r" to R_NS, "w14" to DocxReader.W14, "w15" to DocxReader.W15), "w14 w15"),
            COMMENTS_TYPE,
            DocxReader.REL_COMMENTS,
        )
        return name
    }

    private fun extendedPath(pkg: Pkg, create: Boolean): String? {
        DocxReader.relTarget(pkg.part(DocxReader.DOCUMENT_RELS), DocxReader.REL_COMMENTS_EXTENDED)?.takeIf { pkg.exists(it) }?.let { return it }
        if (!create) return null
        val name = "word/commentsExtended.xml"
        check(!pkg.exists(name)) { "unreadable $name" }
        pkg.add(name, newPart("w15:commentsEx", DocxReader.W15, listOf("w15" to DocxReader.W15), "w15"), COMMENTS_EXT_TYPE, DocxReader.REL_COMMENTS_EXTENDED)
        return name
    }

    /** A paragraph id (8 hex digits below 0x80000000) not used in [taken]; the same inputs give the same id. */
    private fun paraId(seed: Int, taken: MutableSet<String>): String {
        var n = (seed.toLong() * 2654435761L and 0x7FFFFFFFL)
        while (true) {
            val id = String.format("%08X", n.coerceAtLeast(1))
            if (taken.add(id)) return id
            n = (n + 1) and 0x7FFFFFFFL
        }
    }

    private fun usedParaIds(vararg docs: Document?): MutableSet<String> {
        val out = HashSet<String>()
        for (d in docs) {
            val all = d?.getElementsByTagNameNS("*", "p") ?: continue
            for (i in 0 until all.length) (all.item(i) as Element).getAttributeNS(DocxReader.W14, "paraId").takeIf { it.isNotEmpty() }?.let { out.add(it) }
        }
        return out
    }

    private fun byId(doc: Document, tag: String, id: String): List<Element> {
        val all = doc.getElementsByTagNameNS(W, tag)
        return (0 until all.length).map { all.item(it) as Element }.filter { it.getAttributeNS(W, "id") == id }
    }

    private fun referenceRun(document: Document, id: String): Element {
        val r = document.createElementNS(W, "w:r")
        val ref = document.createElementNS(W, "w:commentReference")
        ref.setAttributeNS(W, "w:id", id)
        r.appendChild(ref)
        return r
    }

    /** Puts [node] at character [offset] of [p]: before the text there ([before]) or after the text that ends there. */
    private fun placeAt(p: Element, offset: Int, node: Element, before: Boolean) {
        splitAt(p, offset)
        var pos = 0
        var after: Element? = null
        for (seg in DocxReader.segments(p)) {
            val len = seg.length
            if (len == 0) continue
            if (before && pos == offset) {
                seg.el.parentNode.insertBefore(node, seg.el)
                return
            }
            pos += len
            if (pos == offset) after = seg.el
            if (!before && pos == offset) {
                seg.el.parentNode.insertBefore(node, seg.el.nextSibling)
                return
            }
            if (pos > offset) break
        }
        if (after != null) {
            after.parentNode.insertBefore(node, after.nextSibling)
        } else if (offset == 0) {
            // An empty paragraph (or a start before everything): right after its properties.
            val pPr = child(p, "pPr")
            p.insertBefore(node, pPr?.nextSibling ?: p.firstChild)
        } else {
            p.appendChild(node)
        }
    }

    /**
     * Adds a comment: its text in comments.xml, and in the document a range around
     * [from, to) with its reference mark after. A reply ([Op.CommentAdd.parent]) shares its
     * parent's range and is threaded to it in commentsExtended.
     */
    private fun addComment(document: Document, paragraphs: List<Element>, pkg: Pkg, op: Op.CommentAdd, notes: MutableList<String>) {
        val id = op.id.toString()
        val cPath = commentsPath(pkg, create = true)!!
        val comments = pkg.part(cPath)!!
        check(byId(comments, "comment", id).isEmpty()) { "comment $id already exists" }
        val croot = comments.documentElement
        declare(croot, "w14", DocxReader.W14)
        pkg.touch(cPath)
        val taken = usedParaIds(comments, document)

        val c = comments.createElementNS(W, "w:comment")
        c.setAttributeNS(W, "w:id", id)
        c.setAttributeNS(W, "w:author", op.author.ifBlank { "SNScribe" })
        c.setAttributeNS(W, "w:date", op.date)
        c.setAttributeNS(W, "w:initials", op.initials)
        val lines = op.text.split('\n').ifEmpty { listOf("") }
        lines.forEachIndexed { i, line ->
            val p = comments.createElementNS(W, "w:p")
            p.setAttributeNS(DocxReader.W14, "w14:paraId", paraId(op.id * 31 + i, taken))
            p.setAttributeNS(DocxReader.W14, "w14:textId", "77777777")
            if (i == 0) {
                val r = comments.createElementNS(W, "w:r")
                r.appendChild(comments.createElementNS(W, "w:annotationRef"))
                p.appendChild(r)
            }
            if (line.isNotEmpty()) {
                val r = comments.createElementNS(W, "w:r")
                val t = comments.createElementNS(W, "w:t")
                t.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
                t.textContent = line
                r.appendChild(t)
                p.appendChild(r)
            }
            c.appendChild(p)
        }
        croot.appendChild(c)

        val parent = op.parent?.toString()
        val extPath = extendedPath(pkg, create = parent != null)
        if (extPath != null) {
            val ext = pkg.part(extPath)!!
            pkg.touch(extPath)
            fun entry(para: String, parentPara: String?) {
                val e = ext.createElementNS(DocxReader.W15, "w15:commentEx")
                e.setAttributeNS(DocxReader.W15, "w15:paraId", para)
                parentPara?.let { e.setAttributeNS(DocxReader.W15, "w15:paraIdParent", it) }
                e.setAttributeNS(DocxReader.W15, "w15:done", "0")
                ext.documentElement.appendChild(e)
            }
            val parentPara = parent?.let { pid ->
                val pc = byId(comments, "comment", pid).firstOrNull() ?: error("no comment $pid to reply to")
                DocxReader.lastParaId(pc) ?: paraId(op.id * 31 + 17, taken).also { newId ->
                    elementChildren(pc).last { it.localName == "p" }.setAttributeNS(DocxReader.W14, "w14:paraId", newId)
                }.also { pp ->
                    val has = elementChildren(ext.documentElement).any { it.getAttributeNS(DocxReader.W15, "paraId") == pp }
                    if (!has) entry(pp, null)
                }
            }
            entry(DocxReader.lastParaId(c)!!, parentPara)
        }

        if (parent != null) {
            // Beside the parent's own marks.
            val starts = byId(document, "commentRangeStart", parent)
            val ends = byId(document, "commentRangeEnd", parent)
            val refs = byId(document, "commentReference", parent)
            starts.firstOrNull()?.let { s -> s.parentNode.insertBefore(document.createElementNS(W, "w:commentRangeStart").apply { setAttributeNS(W, "w:id", id) }, s.nextSibling) }
            val endAt = ends.firstOrNull()
            if (endAt != null) {
                val e = document.createElementNS(W, "w:commentRangeEnd").apply { setAttributeNS(W, "w:id", id) }
                val afterRef = refs.firstOrNull()?.parentNode?.takeIf { it.parentNode === endAt.parentNode && it.previousSibling === endAt }
                val anchor = afterRef ?: endAt
                anchor.parentNode.insertBefore(e, anchor.nextSibling)
                e.parentNode.insertBefore(referenceRun(document, id), e.nextSibling)
            } else {
                val refRun = refs.firstOrNull()?.parentNode ?: error("comment $parent has no place in the document")
                refRun.parentNode.insertBefore(referenceRun(document, id), refRun.nextSibling)
            }
            notes.add("reply $id to $parent")
            return
        }
        val first = checkNotNull(paragraphs.getOrNull(op.fromPara)) { "no paragraph ${op.fromPara}" }
        val last = checkNotNull(paragraphs.getOrNull(op.toPara)) { "no paragraph ${op.toPara}" }
        val end = document.createElementNS(W, "w:commentRangeEnd").apply { setAttributeNS(W, "w:id", id) }
        placeAt(last, op.to, end, before = false)
        end.parentNode.insertBefore(referenceRun(document, id), end.nextSibling)
        val start = document.createElementNS(W, "w:commentRangeStart").apply { setAttributeNS(W, "w:id", id) }
        if (first === last && op.from == op.to) end.parentNode.insertBefore(start, end) else placeAt(first, op.from, start, before = true)
        notes.add("comment $id on p${op.fromPara}:${op.from}–p${op.toPara}:${op.to}")
    }

    /** Deletes comments [ids]: their text, their thread entries, and their marks in the document. */
    private fun deleteComments(document: Document, pkg: Pkg, op: Op.CommentDelete, notes: MutableList<String>) {
        val cPath = commentsPath(pkg, create = false) ?: error("the document has no comments")
        val comments = pkg.part(cPath)!!
        val rels = pkg.part(DocxReader.DOCUMENT_RELS)
        val paraIds = HashSet<String>()
        for (id in op.ids.map { it.toString() }) {
            val c = byId(comments, "comment", id).firstOrNull() ?: error("no comment $id")
            elementChildren(c).filter { it.localName == "p" }.forEach { p -> p.getAttributeNS(DocxReader.W14, "paraId").takeIf { it.isNotEmpty() }?.let { paraIds.add(it) } }
            c.parentNode.removeChild(c)
            for (tag in listOf("commentRangeStart", "commentRangeEnd")) byId(document, tag, id).forEach { it.parentNode.removeChild(it) }
            for (ref in byId(document, "commentReference", id)) {
                val run = ref.parentNode as Element
                run.removeChild(ref)
                if (elementChildren(run).all { it.localName == "rPr" }) run.parentNode.removeChild(run)
            }
        }
        pkg.touch(cPath)
        extendedPath(pkg, create = false)?.let { path ->
            val ext = pkg.part(path)!!
            elementChildren(ext.documentElement).filter { it.getAttributeNS(DocxReader.W15, "paraId") in paraIds }.forEach { ext.documentElement.removeChild(it) }
            pkg.touch(path)
        }
        // Word's newer id parts: drop the entries of the deleted comments.
        val durable = HashSet<String>()
        DocxReader.relTarget(rels, REL_COMMENTS_IDS)?.takeIf { pkg.exists(it) }?.let { path ->
            val ids = pkg.part(path)!!
            elementChildren(ids.documentElement).filter { it.getAttributeNS(W16CID, "paraId") in paraIds }.forEach {
                durable.add(it.getAttributeNS(W16CID, "durableId"))
                ids.documentElement.removeChild(it)
            }
            pkg.touch(path)
        }
        DocxReader.relTarget(rels, REL_COMMENTS_EXTENSIBLE)?.takeIf { pkg.exists(it) && durable.isNotEmpty() }?.let { path ->
            val cex = pkg.part(path)!!
            elementChildren(cex.documentElement).filter { it.getAttributeNS(W16CEX, "durableId") in durable }.forEach { cex.documentElement.removeChild(it) }
            pkg.touch(path)
        }
        notes.add("deleted comments ${op.ids}")
    }

    private const val R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private const val REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    private const val REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    private const val REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    private const val HEADER_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    private const val FOOTER_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

    /**
     * Writes the default header or footer. Every default header part the sections use is
     * rewritten (sections without their own inherit the one before, as in Word); a document
     * with none gets a new part on its first section, which the rest inherit. A part holding
     * more than text (a logo, a table) keeps all of it: only its page-number line changes
     * (DocxReader.editableLine), or one is added at the bottom.
     */
    private fun headerFooter(document: Document, pkg: Pkg, op: Op.HeaderFooter, notes: MutableList<String>) {
        val isHeader = op.kind == "header"
        val refTag = if (isHeader) "headerReference" else "footerReference"
        val all = document.getElementsByTagNameNS(W, "sectPr")
        // A sectPr inside w:sectPrChange is a tracked earlier version: it can't hold references.
        val sects = (0 until all.length).map { all.item(it) as Element }.filter { (it.parentNode as? Element)?.localName != "sectPrChange" }
        fun isDefault(e: Element) = e.localName == refTag && e.getAttributeNS(W, "type").let { t -> t.isEmpty() || t == "default" }
        val parts = sects.flatMap { elementChildren(it).filter(::isDefault) }
            .map { it.getAttributeNS(R_NS, "id") }.distinct()
            .mapNotNull { id -> pkg.target(id)?.takeIf { pkg.exists(it) }?.let { id to it } }
        val rId: String
        if (parts.isEmpty()) {
            val name = pkg.freeName(if (isHeader) "word/header" else "word/footer")
            rId = pkg.add(name, fill(isHeader, op), if (isHeader) HEADER_TYPE else FOOTER_TYPE, if (isHeader) REL_HEADER else REL_FOOTER)!!
            notes.add("${op.kind}: added $name")
        } else {
            for ((_, name) in parts) {
                val old = pkg.part(name)!!
                if (!DocxReader.hasOtherContent(old)) {
                    pkg.replace(name, fill(isHeader, op))
                    continue
                }
                val line = old.importNode(fill(isHeader, op).documentElement.firstChild, true)
                val at = DocxReader.editableLine(old)
                if (at != null) old.documentElement.replaceChild(line, at) else old.documentElement.appendChild(line)
                pkg.touch(name)
            }
            rId = parts.first().first
            notes.add("${op.kind}: rewrote ${parts.joinToString { it.second }}")
        }
        // The first section has nothing to inherit from: it needs its own reference.
        sects.firstOrNull()?.takeIf { s -> elementChildren(s).none(::isDefault) }?.let { sect ->
            val ref = document.createElementNS(W, "w:$refTag")
            ref.setAttributeNS(W, "w:type", "default")
            ref.setAttributeNS(R_NS, "r:id", rId)
            // References come first in sectPr: headers, then footers.
            val before = if (isHeader) sect.firstChild else elementChildren(sect).firstOrNull { it.localName != "headerReference" }
            sect.insertBefore(ref, before)
        }
        if (sects.any { s -> child(s, "titlePg") != null }) notes.add("${op.kind}: the document has a different first page; its first page keeps its own")
    }

    /** A header or footer part: one paragraph aligned [op.align], the text, and when asked Word's PAGE field. */
    private fun fill(isHeader: Boolean, op: Op.HeaderFooter): Document {
        val part = javax.xml.parsers.DocumentBuilderFactory.newInstance().apply { isNamespaceAware = true }.newDocumentBuilder().newDocument()
        val root = part.createElementNS(W, if (isHeader) "w:hdr" else "w:ftr")
        root.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", R_NS)
        part.appendChild(root)
        val p = part.createElementNS(W, "w:p")
        root.appendChild(p)
        val pPr = part.createElementNS(W, "w:pPr")
        val jc = part.createElementNS(W, "w:jc")
        jc.setAttributeNS(W, "w:val", if (op.align == "justify") "both" else op.align)
        pPr.appendChild(jc)
        p.appendChild(pPr)
        if (op.text.isNotEmpty()) {
            val r = part.createElementNS(W, "w:r")
            val t = part.createElementNS(W, "w:t")
            t.setAttributeNS(XMLConstants.XML_NS_URI, "xml:space", "preserve")
            t.textContent = if (op.pageNumber) "${op.text.trimEnd()} " else op.text
            r.appendChild(t)
            p.appendChild(r)
        }
        if (op.pageNumber) {
            val fld = part.createElementNS(W, "w:fldSimple")
            fld.setAttributeNS(W, "w:instr", " PAGE ")
            val r = part.createElementNS(W, "w:r")
            val t = part.createElementNS(W, "w:t")
            t.textContent = "1"
            r.appendChild(t)
            fld.appendChild(r)
            p.appendChild(fld)
        }
        return part
    }

    /**
     * Wraps characters [start, end) in a w:hyperlink to [op.url]. The runs covered must sit
     * directly in the paragraph (not already in a link, field or control); the text takes
     * the document's Hyperlink character style, or blue underline when it has none.
     */
    private fun link(document: Document, p: Element, op: Op.Link, pkg: Pkg, styleIds: StyleIds, notes: MutableList<String>) {
        require(op.end > op.start) { "empty link" }
        splitAt(p, op.start)
        splitAt(p, op.end)
        var offset = 0
        val covered = ArrayList<Element>()
        for (seg in DocxReader.segments(p)) {
            val s0 = offset
            offset += seg.length
            if (seg.length == 0 || offset <= op.start || s0 >= op.end) continue
            check(seg.isRun && seg.el.parentNode === p && !seg.locked) { "paragraph ${op.para}: part of that text is already a link, field or control" }
            covered.add(seg.el)
        }
        check(covered.isNotEmpty()) { "nothing to link" }
        // The covered runs must be adjacent: nothing else between them in the paragraph.
        var n: org.w3c.dom.Node? = covered.first()
        for (el in covered.drop(1)) {
            do n = n?.nextSibling while (n != null && n !is Element)
            check(n === el) { "paragraph ${op.para}: something sits inside that text; can't link it" }
        }
        val rId = pkg.relate(REL_HYPERLINK, op.url, external = true)
        val h = document.createElementNS(W, "w:hyperlink")
        h.setAttributeNS(R_NS, "r:id", rId)
        h.setAttributeNS(W, "w:history", "1")
        p.insertBefore(h, covered.first())
        val styleId = styleIds.characterId("hyperlink")
        for (r in covered) {
            h.appendChild(r)
            val rPr = child(r, "rPr") ?: document.createElementNS(W, "w:rPr").also { r.insertBefore(it, r.firstChild) }
            if (styleId != null) {
                child(rPr, "rStyle")?.let { rPr.removeChild(it) }
                val st = document.createElementNS(W, "w:rStyle")
                st.setAttributeNS(W, "w:val", styleId)
                insertInOrder(rPr, st, RPR_ORDER)
            } else {
                for ((tag, value) in listOf("color" to "0563C1", "u" to "single")) {
                    child(rPr, tag)?.let { rPr.removeChild(it) }
                    val e = document.createElementNS(W, "w:$tag")
                    e.setAttributeNS(W, "w:val", value)
                    insertInOrder(rPr, e, RPR_ORDER)
                }
            }
        }
        notes.add("link p${op.para} [${op.start},${op.end}) → ${op.url}")
    }

    /** Unwraps hyperlinks overlapping [start, end): their runs move back into the paragraph. */
    private fun unlink(p: Element, op: Op.Unlink, notes: MutableList<String>) {
        var offset = 0
        val links = LinkedHashSet<Element>()
        for (seg in DocxReader.segments(p)) {
            val s0 = offset
            offset += seg.length
            val parent = seg.el.parentNode
            // A caret (start == end) counts when it touches the link, at either end, as on screen.
            val touches = if (op.start == op.end) s0 <= op.start && offset >= op.start else s0 < op.end && offset > op.start
            if (seg.link && parent is Element && parent.localName == "hyperlink" && seg.length > 0 && touches) links.add(parent)
        }
        for (h in links) {
            while (h.firstChild != null) {
                val c = h.firstChild
                h.removeChild(c)
                h.parentNode.insertBefore(c, h)
                if (c is Element && c.localName == "r") {
                    child(c, "rPr")?.let { rPr ->
                        child(rPr, "rStyle")?.takeIf { it.getAttributeNS(W, "val").equals("Hyperlink", ignoreCase = true) }?.let { rPr.removeChild(it) }
                        child(rPr, "color")?.takeIf { it.getAttributeNS(W, "val").equals("0563C1", ignoreCase = true) }?.let { rPr.removeChild(it) }
                        child(rPr, "u")?.let { rPr.removeChild(it) }
                    }
                }
            }
            h.parentNode.removeChild(h)
        }
        notes.add("unlink p${op.para}: ${links.size} link(s)")
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
            // Off is written out ("0"), so it also removes a page break the paragraph's style gives it.
            val el = document.createElementNS(W, "w:pageBreakBefore")
            if (!on) el.setAttributeNS(W, "w:val", "0")
            insertInOrder(pPr, el, PPR_ORDER)
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

    /** Word's highlight colours (ST_HighlightColor) DOCX offers. */
    private val HIGHLIGHTS = setOf("yellow", "green", "blue", "magenta", "red", "cyan")

    /** Sets a run property explicitly on or off, so styles cannot override the user's choice. */
    private fun setRunProperty(document: Document, r: Element, prop: String, on: Boolean, colour: String? = null) {
        val (name, value) = when (prop) {
            "b" -> "b" to (if (on) null else "0")
            "i" -> "i" to (if (on) null else "0")
            "u" -> "u" to (if (on) "single" else "none")
            "s" -> "strike" to (if (on) null else "0")
            "h" -> "highlight" to (if (on) (colour?.takeIf { it in HIGHLIGHTS } ?: "yellow") else "none")
            // Superscript and subscript share w:vertAlign: turning one on replaces the other.
            "sup" -> "vertAlign" to (if (on) "superscript" else "baseline")
            "sub" -> "vertAlign" to (if (on) "subscript" else "baseline")
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
        private val characterByName = HashMap<String, String>()
        private var defaultId: String? = null

        init {
            styles?.documentElement?.let { root ->
                for (s in elementChildren(root).filter { it.localName == "style" && it.getAttributeNS(W, "type") == "character" }) {
                    child(s, "name")?.getAttributeNS(W, "val")?.lowercase()?.let { characterByName.putIfAbsent(it, s.getAttributeNS(W, "styleId")) }
                }
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

        /**
         * Rewrites the look of the style for [def].kind (creating it first when missing):
         * run font/size/bold/italic, and alignment, no indent, and line spacing — the rest of
         * the style (outline level, keep-with-next) stays. Colour is removed: paper formats
         * want black headings.
         */
        fun redefine(def: StyleDef) {
            val doc = styles ?: return
            val id = idFor(def.kind) ?: return
            val s = raw[id] ?: elementChildren(doc.documentElement).firstOrNull {
                it.localName == "style" && it.getAttributeNS(W, "styleId") == id
            }?.also { raw[id] = it } ?: return
            fun el(parent: Element, tag: String, order: List<String>): Element =
                child(parent, tag) ?: doc.createElementNS(W, "w:$tag").also { insertInOrder(parent, it, order) }
            val styleOrder = listOf("name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
                "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply",
                "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr")
            val pPr = el(s, "pPr", styleOrder)
            val rPr = el(s, "rPr", styleOrder)
            def.align?.let { a -> el(pPr, "jc", PPR_ORDER).setAttributeNS(W, "w:val", if (a == "justify") "both" else a) }
            // No indent of its own (a heading based on List Paragraph would inherit one otherwise).
            el(pPr, "ind", PPR_ORDER).apply {
                listOf("left", "start", "right", "end", "hanging", "firstLine").forEach { removeAttributeNS(W, it) }
                setAttributeNS(W, "w:left", "0")
                setAttributeNS(W, "w:firstLine", "0")
            }
            def.line?.let { line ->
                el(pPr, "spacing", PPR_ORDER).apply {
                    setAttributeNS(W, "w:before", "0")
                    setAttributeNS(W, "w:after", "0")
                    setAttributeNS(W, "w:line", line.toString())
                    setAttributeNS(W, "w:lineRule", "auto")
                }
            }
            def.font?.let { f ->
                el(rPr, "rFonts", RPR_ORDER).apply {
                    listOf("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme").forEach { removeAttributeNS(W, it) }
                    setAttributeNS(W, "w:ascii", f)
                    setAttributeNS(W, "w:hAnsi", f)
                    setAttributeNS(W, "w:cs", f)
                }
            }
            fun onOff(tag: String, on: Boolean?) {
                on ?: return
                el(rPr, tag, RPR_ORDER).apply { if (on) removeAttributeNS(W, "val") else setAttributeNS(W, "w:val", "0") }
            }
            onOff("b", def.bold)
            onOff("bCs", def.bold)
            onOff("i", def.italic)
            onOff("iCs", def.italic)
            def.size?.let { sz ->
                el(rPr, "sz", RPR_ORDER).setAttributeNS(W, "w:val", sz.toString())
                el(rPr, "szCs", RPR_ORDER).setAttributeNS(W, "w:val", sz.toString())
            }
            child(rPr, "color")?.let { rPr.removeChild(it) }
            listOf("caps", "smallCaps").forEach { t -> child(rPr, t)?.let { rPr.removeChild(it) } }
            added = true
        }

        /** A character style's id by its built-in name ("hyperlink"), or null. */
        fun characterId(name: String): String? = characterByName[name]

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
                "heading3" -> Quad("heading 3", "Heading3", 2, 24)
                "title" -> Quad("title", "Title", null, 56)
                "quote" -> Quad("quote", "Quote", null, 0)
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
            if (kind == "quote") {
                // A block quotation: indented both sides, italic, the body size.
                val spacing = add(pPr, "spacing")
                spacing.setAttributeNS(W, "w:before", "120")
                spacing.setAttributeNS(W, "w:after", "120")
                val ind = add(pPr, "ind")
                ind.setAttributeNS(W, "w:left", "720")
                ind.setAttributeNS(W, "w:right", "720")
                add(add(s, "rPr"), "i")
            } else {
                add(pPr, "keepNext")
                val spacing = add(pPr, "spacing")
                spacing.setAttributeNS(W, "w:before", "240")
                spacing.setAttributeNS(W, "w:after", "120")
                if (outline != null) add(pPr, "outlineLvl", outline.toString())
                val rPr = add(s, "rPr")
                add(rPr, "b")
                add(rPr, "sz", size.toString())
            }
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
