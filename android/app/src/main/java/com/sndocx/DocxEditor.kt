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
 * Formatting never changes text, so ops can be applied in any order, and the saved copy is
 * verified by re-reading it: every paragraph's text must equal the original's.
 */
object DocxEditor {
    sealed class Op {
        /** prop: "b", "i", "u" or "h". */
        data class Format(val para: Int, val start: Int, val end: Int, val prop: String, val on: Boolean) : Op()

        /** kind: "heading1", "heading2", "title" or "normal". */
        data class Style(val para: Int, val kind: String) : Op()
    }

    data class Saved(val dest: File, val changedParts: List<String>, val notes: List<String>)

    private const val DOCUMENT_PART = "word/document.xml"
    private const val STYLES_PART = "word/styles.xml"

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
    fun save(src: File, ops: List<Op>, dest: File, workDir: File): Saved {
        require(src.length() <= DocxReader.MAX_DOCX_BYTES) { "too large: ${src.length()} bytes" }
        val notes = ArrayList<String>()
        val changed = ArrayList<String>()
        workDir.mkdirs()
        val temp = File(workDir, "saving-${System.nanoTime()}.docx")
        try {
            ZipFile(src).use { zip ->
                val documentXml = DocxReader.readEntry(zip, zip.getEntry(DOCUMENT_PART) ?: error("no $DOCUMENT_PART"))
                val stylesEntry = zip.getEntry(STYLES_PART)
                val document = DocxReader.parse(documentXml)
                val styles = stylesEntry?.let { DocxReader.parse(DocxReader.readEntry(zip, it)) }
                val stylesChanged = apply(document, styles, ops, notes)
                val newDocument = serialize(document)
                val newStyles = if (stylesChanged && styles != null) serialize(styles) else null

                ZipOutputStream(FileOutputStream(temp)).use { out ->
                    for (entry in zip.entries()) {
                        DocxReader.checkEntryName(entry.name)
                        val bytes = when {
                            entry.name == DOCUMENT_PART && ops.isNotEmpty() -> newDocument.also { changed.add(entry.name) }
                            entry.name == STYLES_PART && newStyles != null -> newStyles.also { changed.add(entry.name) }
                            else -> DocxReader.readEntry(zip, entry)
                        }
                        putEntry(out, entry, bytes)
                    }
                }
            }
            verify(src, temp, changed)
            copy(temp, dest)
            return Saved(dest, changed, notes)
        } finally {
            temp.delete()
        }
    }

    /** Every paragraph's text unchanged, every untouched part byte-identical. */
    private fun verify(src: File, written: File, changed: List<String>) {
        val before = DocxReader.read(src).blocks.filterIsInstance<DocxReader.Paragraph>()
        val after = DocxReader.read(written).blocks.filterIsInstance<DocxReader.Paragraph>()
        check(before.size == after.size) { "verify: paragraph count ${before.size} → ${after.size}" }
        for ((a, b) in before.zip(after)) {
            check(a.text == b.text) { "verify: text of paragraph ${a.index} changed" }
        }
        val ha = hashes(src)
        val hb = hashes(written)
        check(ha.keys == hb.keys) { "verify: package parts differ: ${ha.keys - hb.keys} / ${hb.keys - ha.keys}" }
        for ((name, hash) in ha) {
            if (name !in changed) check(hb[name] == hash) { "verify: $name changed but was not edited" }
        }
    }

    // ---------------------------------------------------------------- apply

    /** Applies [ops] to the DOM. Returns whether styles.xml was changed (a style was added). */
    fun apply(document: Document, styles: Document?, ops: List<Op>, notes: MutableList<String>): Boolean {
        val paragraphs = bodyParagraphs(document)
        val styleIds = StyleIds(styles)
        for (op in ops) {
            val p = paragraphs.getOrNull(
                when (op) {
                    is Op.Format -> op.para
                    is Op.Style -> op.para
                },
            )
            if (p == null) {
                notes.add("skipped $op: no such paragraph")
                continue
            }
            when (op) {
                is Op.Format -> format(document, p, op, notes)
                is Op.Style -> setStyle(document, p, styleIds.idFor(op.kind))
            }
        }
        return styleIds.added
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

        init {
            styles?.documentElement?.let { root ->
                for (s in elementChildren(root).filter { it.localName == "style" && it.getAttributeNS(W, "type") == "paragraph" }) {
                    val id = s.getAttributeNS(W, "styleId")
                    ids.add(id)
                    child(s, "name")?.getAttributeNS(W, "val")?.lowercase()?.let { byName.putIfAbsent(it, id) }
                }
            }
        }

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
