package com.sndocx

import com.facebook.react.bridge.Arguments
import com.facebook.react.bridge.Promise
import com.facebook.react.bridge.ReactApplicationContext
import com.facebook.react.bridge.ReactContextBaseJavaModule
import com.facebook.react.bridge.ReactMethod
import com.facebook.react.bridge.ReadableArray
import com.facebook.react.bridge.WritableMap
import android.graphics.Typeface
import com.facebook.react.common.assets.ReactFontManager
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale
import java.util.concurrent.Executors
import java.util.concurrent.TimeUnit

/**
 * The bridge to [DocxReader], plus the session log.
 *
 * Log: one append-only file per PluginHost process, EXPORT/sn-docx-log-<date>-<time>.txt.
 * Writes to EXPORT are refused until plugin.permission.FILE:WRITE is granted (the probe lost
 * two logs that way), so log() reports a refusal instead of swallowing it.
 */
class DocxModule(private val reactContext: ReactApplicationContext) : ReactContextBaseJavaModule(reactContext) {

    override fun getName() = "Docx"

    private val worker = Executors.newSingleThreadExecutor()
    private val timer = Executors.newSingleThreadScheduledExecutor()

    companion object {
        /** Bumped with each native change; first log line, to spot stale installs. */
        const val NATIVE_BUILD = 22
        private val EXPORT_DIR = File("/storage/emulated/0/EXPORT")
        private const val LOG_MAX_BYTES = 2L * 1024 * 1024
        private const val LOG_LINE_MAX = 4000

        @Volatile private var logFile: File? = null
        @Volatile private var capped = false

        /** Null when written, else why not. Never throws. */
        @Synchronized
        fun appendLog(line: String): String? {
            try {
                if (!EXPORT_DIR.isDirectory) EXPORT_DIR.mkdirs()
                val file = logFile ?: File(
                    EXPORT_DIR,
                    "sn-docx-log-" + SimpleDateFormat("yyyyMMdd-HHmmss", Locale.US).format(Date()) + ".txt",
                ).also {
                    write(it, "native $NATIVE_BUILD: new log ${it.name}")
                    logFile = it
                }
                if (capped) return "log cap reached"
                if (file.length() >= LOG_MAX_BYTES) {
                    capped = true
                    write(file, "LOG CAP REACHED — later lines are not recorded")
                    return "log cap reached"
                }
                write(file, if (line.length > LOG_LINE_MAX) line.take(LOG_LINE_MAX) + " …(+${line.length - LOG_LINE_MAX})" else line)
                return null
            } catch (t: Throwable) {
                return t.toString()
            }
        }

        private fun write(file: File, line: String) {
            val stamp = SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS", Locale.US).format(Date())
            FileOutputStream(file, true).use { it.write("$stamp  $line\n".toByteArray(Charsets.UTF_8)) }
        }
    }

    override fun getConstants(): Map<String, Any> = mapOf("NATIVE_BUILD" to NATIVE_BUILD)

    override fun invalidate() {
        worker.shutdown()
        timer.shutdown()
        super.invalidate()
    }

    @ReactMethod
    fun log(line: String, promise: Promise) {
        timer.execute { promise.resolve(appendLog(line)) }
    }

    @ReactMethod
    fun logName(promise: Promise) {
        timer.execute {
            appendLog("panel opened")
            promise.resolve(logFile?.name ?: "(not started)")
        }
    }

    /** Native wait: JS timers stop while the panel is hidden. */
    @ReactMethod
    fun delay(ms: Double, promise: Promise) {
        timer.schedule({ promise.resolve(null) }, ms.toLong(), TimeUnit.MILLISECONDS)
    }

    /**
     * Opens a document for reading. Resolves {path, name, bytes, ms, report, blocks}; see
     * src/model/docx.ts for the shape. Rejects with the reason when the file can't be read.
     */
    @ReactMethod
    fun open(path: String, promise: Promise) {
        worker.execute {
            val t0 = System.currentTimeMillis()
            try {
                val file = File(path)
                val result = DocxReader.read(file)
                // The document's pictures, for the screen (also used by preview of this document).
                pictures = runCatching {
                    DocxReader.extractImages(file, File(reactContext.cacheDir, "sn-docx/pictures/" + Integer.toHexString(file.path.hashCode())))
                }.onFailure { appendLog("pictures not read: $it") }.getOrNull() ?: emptyMap()
                val ms = System.currentTimeMillis() - t0
                appendLog("open ${file.name}: ${file.length()} B, ${result.blocks.size} blocks, $ms ms, ${result.report}")
                promise.resolve(Arguments.createMap().apply {
                    putString("path", file.path)
                    putString("name", file.name)
                    putDouble("bytes", file.length().toDouble())
                    putDouble("ms", ms.toDouble())
                    putMap("report", report(result.report))
                    // Handwritten margin notes: their pictures, for the margin column.
                    runCatching {
                        DocxReader.extractInk(file, File(reactContext.filesDir, "sn-docx/ink/" + Integer.toHexString(file.path.hashCode())))
                    }.onFailure { appendLog("ink pictures not read: $it") }.getOrNull()?.takeIf { it.isNotEmpty() }?.let { inks ->
                        putMap("inks", Arguments.createMap().apply { for ((k, v) in inks) putString(k, v) })
                    }
                    if (result.otherRevisions > 0) putInt("otherRevisions", result.otherRevisions)
                    putMap("looks", Arguments.createMap().apply {
                        for ((kind, l) in result.looks) putMap(kind, Arguments.createMap().apply {
                            l.font?.let { putString("bf", it) }
                            l.size?.let { putInt("bs", it) }
                            if (l.bold) putBoolean("sb", true)
                            if (l.italic) putBoolean("si", true)
                            putString("align", l.align)
                            putInt("indent", l.indent)
                        })
                    })
                    putArray("comments", Arguments.createArray().apply {
                        for (c in result.comments) pushMap(Arguments.createMap().apply {
                            putString("id", c.id)
                            putString("author", c.author)
                            putString("initials", c.initials)
                            putString("date", c.date)
                            putString("text", c.text)
                            c.parent?.let { putString("parent", it) }
                            if (c.done) putBoolean("done", true)
                            if (c.pictures > 0) putInt("pictures", c.pictures)
                        })
                    })
                    putArray("footnotes", Arguments.createArray().apply {
                        for (f in result.footnotes) pushMap(Arguments.createMap().apply {
                            putString("id", f.id)
                            putArray("pieces", notePieces(f.pieces))
                        })
                    })
                    result.page?.let { pg ->
                        putMap("page", Arguments.createMap().apply {
                            putInt("width", pg.width); putInt("height", pg.height)
                            putInt("top", pg.top); putInt("right", pg.right); putInt("bottom", pg.bottom); putInt("left", pg.left)
                            putBoolean("landscape", pg.landscape)
                        })
                    }
                    for ((key, hf) in listOf("header" to result.header, "footer" to result.footer)) {
                        hf?.let {
                            putMap(key, Arguments.createMap().apply {
                                putString("text", it.text)
                                putBoolean("pageNumber", it.pageNumber)
                                putString("align", it.align)
                                putBoolean("other", it.other)
                            })
                        }
                    }
                    putArray("blocks", Arguments.createArray().apply { result.blocks.forEach { pushMap(block(it)) } })
                    putMap("lists", Arguments.createMap().apply {
                        for ((id, def) in result.lists) putMap(id.toString(), Arguments.createMap().apply {
                            putArray("levels", Arguments.createArray().apply {
                                for (l in def.levels) {
                                    if (l == null) pushNull() else pushMap(Arguments.createMap().apply {
                                        putString("fmt", l.fmt)
                                        putString("text", l.text)
                                        putInt("start", l.start)
                                    })
                                }
                            })
                            putMap("starts", Arguments.createMap().apply { for ((k, v) in def.starts) putInt(k.toString(), v) })
                        })
                    })
                })
            } catch (t: Throwable) {
                appendLog("open FAILED $path: $t")
                promise.reject("DOCX_OPEN_FAILED", t.message ?: t.toString(), t)
            }
        }
    }

    /**
     * Saves a copy with [ops] applied. [dest] empty = a new `<name>-edited.docx` beside the
     * original; otherwise that copy (from an earlier save this session) is overwritten.
     * Built and verified in private storage first (DocxEditor.save). Resolves
     * {dest, name, ms, changed, notes}; rejects with the reason, leaving any file untouched.
     */
    @ReactMethod
    fun save(srcPath: String, ops: ReadableArray, dest: String, expected: ReadableArray, key: String, promise: Promise) {
        worker.execute {
            val t0 = System.currentTimeMillis()
            try {
                val src = File(srcPath)
                val parsed = parseOps(ops)
                val target = if (dest.isEmpty()) DocxEditor.editedCopyName(src) else File(dest)
                val expect = (0 until expected.size()).map { expected.getString(it) ?: "" }
                // A save that fails while writing keeps its checked file as the newest previous version.
                val keep: ((File) -> File?)? = if (key.isEmpty()) null else { verified ->
                    File(home("backups/" + safeName(key)), "${System.currentTimeMillis()}.docx").also { verified.copyTo(it, overwrite = true) }
                }
                val saved = DocxEditor.save(src, parsed, target, File(reactContext.cacheDir, "saving"), expect, keep)
                val ms = System.currentTimeMillis() - t0
                appendLog("save ${src.name} → ${target.path}: ${parsed.size} ops, ${saved.changedParts}, $ms ms")
                saved.notes.forEach { appendLog("  $it") }
                promise.resolve(Arguments.createMap().apply {
                    putString("dest", target.path)
                    putString("name", target.name)
                    putDouble("ms", ms.toDouble())
                    putArray("changed", Arguments.fromList(saved.changedParts))
                })
            } catch (t: Throwable) {
                appendLog("save FAILED $srcPath: $t")
                promise.reject("DOCX_SAVE_FAILED", t.message ?: t.toString(), t)
            }
        }
    }

    /** A footnote's text from JS: [{t, i?, b?}]. */
    private fun pieces(a: ReadableArray?): List<DocxReader.NotePiece> = (0 until (a?.size() ?: 0)).mapNotNull { a!!.getMap(it) }.map {
        DocxReader.NotePiece(it.getString("t") ?: "", it.hasKey("i") && it.getBoolean("i"), it.hasKey("b") && it.getBoolean("b"))
    }

    private fun parseOps(ops: ReadableArray): List<DocxEditor.Op> = (0 until ops.size()).mapNotNull { i -> ops.getMap(i) }.map { m ->
        when (m.getString("op")) {
            "format" -> DocxEditor.Op.Format(
                m.getInt("para"), m.getInt("start"), m.getInt("end"),
                m.getString("prop") ?: "", m.getBoolean("on"),
                if (m.hasKey("value") && !m.isNull("value")) m.getString("value") else null,
            )
            "style" -> DocxEditor.Op.Style(m.getInt("para"), m.getString("kind") ?: "normal")
            "text" -> DocxEditor.Op.Text(
                m.getInt("para"), m.getInt("start"), m.getInt("end"), m.getString("text") ?: "",
            )
            "split" -> DocxEditor.Op.Split(m.getInt("para"), m.getInt("offset"))
            "join" -> DocxEditor.Op.Join(m.getInt("para"))
            "runStyle" -> DocxEditor.Op.RunStyle(
                m.getInt("para"), m.getInt("start"), m.getInt("end"),
                if (m.hasKey("font") && !m.isNull("font")) m.getString("font") else null,
                if (m.hasKey("size") && !m.isNull("size")) m.getInt("size") else null,
            )
            "para" -> {
                fun int(k: String) = if (m.hasKey(k) && !m.isNull(k)) m.getInt(k) else null
                fun str(k: String) = if (m.hasKey(k) && !m.isNull(k)) m.getString(k) else null
                DocxEditor.Op.ParaProps(
                    para = m.getInt("para"),
                    align = str("align"),
                    line = int("line"),
                    lineRule = str("lineRule"),
                    before = int("before"),
                    after = int("after"),
                    first = int("first"),
                    pageBreakBefore = if (m.hasKey("pb") && !m.isNull("pb")) m.getBoolean("pb") else null,
                )
            }
            "page" -> {
                fun int(k: String) = if (m.hasKey(k) && !m.isNull(k)) m.getInt(k) else null
                DocxEditor.Op.PageSetup(
                    width = int("width"), height = int("height"),
                    landscape = if (m.hasKey("landscape") && !m.isNull("landscape")) m.getBoolean("landscape") else null,
                    top = int("top"), right = int("right"), bottom = int("bottom"), left = int("left"),
                )
            }
            "defaults" -> DocxEditor.Op.Defaults(
                if (m.hasKey("font") && !m.isNull("font")) m.getString("font") else null,
                if (m.hasKey("size") && !m.isNull("size")) m.getInt("size") else null,
            )
            "headerFooter" -> DocxEditor.Op.HeaderFooter(
                m.getString("kind") ?: "header", m.getString("text") ?: "",
                m.hasKey("pageNumber") && m.getBoolean("pageNumber"), m.getString("align") ?: "right",
            )
            "link" -> DocxEditor.Op.Link(m.getInt("para"), m.getInt("start"), m.getInt("end"), m.getString("url") ?: "")
            "unlink" -> DocxEditor.Op.Unlink(m.getInt("para"), m.getInt("start"), m.getInt("end"))
            "styleDefs" -> DocxEditor.Op.StyleDefs(m.getArray("defs")?.let { arr ->
            (0 until arr.size()).mapNotNull { arr.getMap(it) }.map { d ->
                fun str(k: String) = if (d.hasKey(k) && !d.isNull(k)) d.getString(k) else null
                fun int(k: String) = if (d.hasKey(k) && !d.isNull(k)) d.getInt(k) else null
                fun bool(k: String) = if (d.hasKey(k) && !d.isNull(k)) d.getBoolean(k) else null
                DocxEditor.StyleDef(str("kind") ?: "", str("font"), int("size"), bool("bold"), bool("italic"), str("align"), int("line"))
            }
        }.orEmpty())
        "comment" -> DocxEditor.Op.CommentAdd(
            id = m.getInt("id"),
            fromPara = m.getInt("fromPara"), from = m.getInt("from"),
            toPara = m.getInt("toPara"), to = m.getInt("to"),
            text = m.getString("text") ?: "",
            author = m.getString("author") ?: "",
            initials = m.getString("initials") ?: "",
            date = m.getString("date") ?: "",
            parent = if (m.hasKey("parent") && !m.isNull("parent")) m.getInt("parent") else null,
        )
        "uncomment" -> DocxEditor.Op.CommentDelete(m.getArray("ids")?.let { a -> (0 until a.size()).map { a.getInt(it) } }.orEmpty())
        "tableCell" -> DocxEditor.Op.TableCell(m.getInt("table"), m.getInt("row"), m.getInt("cell"), pieces(m.getArray("pieces")))
        "tableRowAdd" -> DocxEditor.Op.TableRowAdd(m.getInt("table"), m.getInt("row"), m.getBoolean("below"))
        "tableRowDelete" -> DocxEditor.Op.TableRowDelete(m.getInt("table"), m.getInt("row"))
        "tableInsert" -> DocxEditor.Op.TableInsert(m.getInt("before"), m.getInt("rows"), m.getInt("cols"), m.hasKey("header") && m.getBoolean("header"), if (m.hasKey("pct")) m.getInt("pct") else 100)
        "tableWidth" -> DocxEditor.Op.TableWidth(m.getInt("table"), m.getInt("pct"))
        "tableDelete" -> DocxEditor.Op.TableDelete(m.getInt("table"))
        "imageSize" -> DocxEditor.Op.ImageSize(m.getInt("para"), m.getInt("at"), m.getDouble("cx").toLong(), m.getDouble("cy").toLong())
        "imageDelete" -> DocxEditor.Op.ImageDelete(m.getInt("para"), m.getInt("at"))
        "footnote" -> DocxEditor.Op.FootnoteAdd(m.getInt("para"), m.getInt("at"), m.getInt("id"), pieces(m.getArray("pieces")))
        "footnoteSet" -> DocxEditor.Op.FootnoteSet(m.getInt("id"), pieces(m.getArray("pieces")))
        "footnoteDelete" -> DocxEditor.Op.FootnoteDelete(m.getInt("id"))
        "image" -> DocxEditor.Op.ImageAdd(m.getInt("para"), m.getInt("at"), m.getString("path") ?: "", m.getDouble("cx").toLong(), m.getDouble("cy").toLong(), if (m.hasKey("alt")) m.getString("alt") ?: "" else "")
        "ink" -> DocxEditor.Op.InkAdd(m.getInt("para"), m.getInt("at"), m.getString("id") ?: "", m.getString("png") ?: "", m.getInt("width"), m.getInt("height"))
        "inkDelete" -> DocxEditor.Op.InkDelete(m.getInt("para"), m.getString("id") ?: "")
        "revision" -> DocxEditor.Op.Revision(m.getInt("para"), m.getString("id") ?: "", m.getBoolean("accept"))
        "list" -> DocxEditor.Op.ListItem(m.getInt("para"), m.getString("kind") ?: "none", m.getString("listId") ?: "")
        "listLevel" -> DocxEditor.Op.ListLevel(m.getInt("para"), m.getInt("delta"))
            else -> throw IllegalArgumentException("unknown op ${m.getString("op")}")
        }
    }

    /**
     * The paragraphs [paras] (empty = all) as they will be once [ops] are applied to
     * [srcPath] — for edits whose result only the file knows exactly (accepting or rejecting
     * tracked changes). Nothing is written. Resolves {blocks: [paragraph blocks]}.
     */
    @ReactMethod
    fun preview(srcPath: String, ops: ReadableArray, paras: ReadableArray, promise: Promise) {
        worker.execute {
            val t0 = System.currentTimeMillis()
            try {
                val parsed = parseOps(ops)
                val result = java.util.zip.ZipFile(File(srcPath)).use { zip ->
                    val pkg = DocxEditor.Pkg(zip)
                    val document = pkg.part(DocxReader.DOCUMENT_PART) ?: error("no document")
                    val styles = pkg.part("word/styles.xml")
                    val rels = pkg.part(DocxReader.DOCUMENT_RELS)
                    val numbering = pkg.part(DocxReader.numberingPart(rels))
                    DocxEditor.apply(document, styles, parsed, ArrayList(), DocxEditor.Lists(numbering), pkg)
                    DocxReader.read(document, styles, numbering, pkg.part(DocxReader.themePart(rels)))
                }
                val want = (0 until paras.size()).map { paras.getInt(it) }.toSet()
                val blocks = result.blocks.filterIsInstance<DocxReader.Paragraph>().filter { want.isEmpty() || it.index in want }
                appendLog("preview ${parsed.size} ops → ${blocks.size} paragraphs, ${System.currentTimeMillis() - t0} ms")
                promise.resolve(Arguments.createMap().apply {
                    putArray("blocks", Arguments.createArray().apply { blocks.forEach { pushMap(block(it)) } })
                })
            } catch (t: Throwable) {
                appendLog("preview FAILED: $t")
                promise.reject("DOCX_PREVIEW_FAILED", t.message ?: t.toString(), t)
            }
        }
    }

    // ---------------------------------------------------------------- new documents

    /**
     * A new blank document named [name] in the Document folder (never over an existing
     * file). Also keeps a pristine copy in private storage: edits are always applied to
     * that copy and saved over the new file, so every save starts from the same blank.
     * Resolves {path, source}.
     */
    @ReactMethod
    fun create(name: String, folder: String, promise: Promise) {
        worker.execute {
            try {
                val dir = File(folder.ifEmpty { "/storage/emulated/0/Document" }).also { if (!it.isDirectory) it.mkdirs() }
                check(dir.isDirectory) { "No such folder: ${dir.path}" }
                val dest = DocxBlank.freeName(dir, name)
                val source = File(File(reactContext.filesDir, "sn-docx-new").also { it.mkdirs() }, "blank-${System.currentTimeMillis()}.docx")
                DocxBlank.write(source)
                source.inputStream().use { input -> FileOutputStream(dest).use { input.copyTo(it) } }
                appendLog("created ${dest.path} (source ${source.name})")
                promise.resolve(Arguments.createMap().apply {
                    putString("path", dest.path)
                    putString("source", source.path)
                })
            } catch (t: Throwable) {
                appendLog("create FAILED: $t")
                promise.reject("DOCX_CREATE_FAILED", t.message ?: t.toString(), t)
            }
        }
    }

    /**
     * A new document named [name] in [folder], copied from the template or document at
     * [templatePath] (.docx or .dotx; a .dotx is turned into a document). Like [create], a
     * pristine copy stays in private storage as the source saves start from.
     * Resolves {path, source}.
     */
    @ReactMethod
    fun createFrom(templatePath: String, name: String, folder: String, promise: Promise) {
        worker.execute {
            try {
                val template = File(templatePath)
                check(template.isFile) { "No such file: ${template.path}" }
                val ext = template.extension.lowercase()
                check(ext == "docx" || ext == "dotx") { "Choose a .docx or .dotx file" }
                val dir = File(folder.ifEmpty { "/storage/emulated/0/Document" }).also { if (!it.isDirectory) it.mkdirs() }
                check(dir.isDirectory) { "No such folder: ${dir.path}" }
                val dest = DocxBlank.freeName(dir, name)
                val source = File(File(reactContext.filesDir, "sn-docx-new").also { it.mkdirs() }, "template-${System.currentTimeMillis()}.docx")
                DocxBlank.fromTemplate(template, source)
                DocxReader.read(source) // refuse anything the reader can't open, before it lands in the folder
                source.inputStream().use { input -> FileOutputStream(dest).use { input.copyTo(it) } }
                appendLog("created ${dest.path} from ${template.path} (source ${source.name})")
                promise.resolve(Arguments.createMap().apply {
                    putString("path", dest.path)
                    putString("source", source.path)
                })
            } catch (t: Throwable) {
                appendLog("createFrom FAILED: $t")
                promise.reject("DOCX_CREATE_FAILED", t.message ?: t.toString(), t)
            }
        }
    }

    /**
     * The folders directly inside [path], for choosing where a new document goes. There is no
     * folder picker in the plugin SDK, so DOCX lists them itself (as SNFolio does). Resolves
     * {folders: [names]} or {error} when the folder can't be listed (scoped storage can
     * refuse it); never rejects.
     */
    @ReactMethod
    fun listFolders(path: String, promise: Promise) {
        worker.execute {
            val result = Arguments.createMap()
            val children = runCatching { File(path).listFiles() }.getOrNull()
            if (children == null) {
                result.putString("error", "This folder can't be listed")
            } else {
                result.putArray("folders", Arguments.fromList(
                    children.filter { it.isDirectory && !it.name.startsWith(".") }.map { it.name }.sortedBy { it.lowercase() },
                ))
            }
            promise.resolve(result)
        }
    }

    // ---------------------------------------------------------------- private storage, versions

    /** DOCX's own folder in PluginHost's private storage: recovery, settings, snapshots, backups. */
    private fun home(sub: String): File = File(File(reactContext.filesDir, "sn-docx"), sub).also { it.mkdirs() }

    private fun safeName(name: String) = name.replace(Regex("[^A-Za-z0-9._-]"), "_").take(80)

    /** Keeps [json] under [name] (recovery records, recent documents, settings). */
    @ReactMethod
    fun store(name: String, json: String, promise: Promise) {
        worker.execute {
            // All or nothing: written beside it, then renamed over it, so an interrupted write
            // never leaves a half record where the last good one was (audit 2026-09-28).
            val dest = File(home("data"), safeName(name) + ".json")
            val temp = File(dest.parentFile, dest.name + ".tmp")
            val r = runCatching {
                temp.writeText(json)
                if (!temp.renameTo(dest)) {
                    dest.delete()
                    check(temp.renameTo(dest)) { "could not replace ${dest.name}" }
                }
            }
            if (r.isFailure) {
                temp.delete()
                appendLog("store $name FAILED: ${r.exceptionOrNull()}")
            }
            promise.resolve(r.exceptionOrNull()?.toString())
        }
    }

    /**
     * Keeps a private copy of [source] (the document as it was opened) under [name], for a
     * recovery record that replays every edit since opening. Resolves the copy's path.
     */
    @ReactMethod
    fun keepBase(source: String, name: String, promise: Promise) {
        worker.execute {
            try {
                val dest = File(home("recovery"), safeName(name) + ".docx")
                val temp = File(dest.parentFile, dest.name + ".tmp")
                File(source).inputStream().use { input -> FileOutputStream(temp).use { input.copyTo(it) } }
                if (!temp.renameTo(dest)) {
                    dest.delete()
                    check(temp.renameTo(dest)) { "could not replace ${dest.name}" }
                }
                promise.resolve(dest.path)
            } catch (t: Throwable) {
                appendLog("keepBase FAILED $source: $t")
                promise.reject("DOCX_BASE", t.message ?: t.toString(), t)
            }
        }
    }

    /** What [store] kept under [name], or null. */
    @ReactMethod
    fun load(name: String, promise: Promise) {
        worker.execute {
            promise.resolve(runCatching { File(home("data"), safeName(name) + ".json").takeIf { it.isFile }?.readText() }.getOrNull())
        }
    }

    /** Forgets [name] (DOCX's own private file; deleting there is allowed). */
    @ReactMethod
    fun forget(name: String, promise: Promise) {
        worker.execute {
            // A recovery record's kept copy of the opened document goes with it.
            runCatching { File(home("recovery"), safeName(name) + ".docx").delete() }
            promise.resolve(runCatching { File(home("data"), safeName(name) + ".json").delete() }.getOrDefault(false))
        }
    }

    /**
     * The words in [words] that the [lang] dictionary does not know (loaded once, on the
     * worker thread; an empty list just loads it). Resolves the misspelled ones.
     */
    @ReactMethod
    fun spellCheck(lang: String, words: ReadableArray, promise: Promise) {
        worker.execute {
            try {
                val t0 = System.currentTimeMillis()
                val sc = SpellChecker.builtIn(lang)
                val list = (0 until words.size()).mapNotNull { words.getString(it) }
                val bad = list.filter { !sc.check(it) }
                if (list.isEmpty()) appendLog("spelling: $lang ready in ${System.currentTimeMillis() - t0} ms")
                promise.resolve(Arguments.fromList(bad))
            } catch (t: Throwable) {
                appendLog("spelling FAILED: $t")
                promise.reject("DOCX_SPELL", t.message ?: t.toString(), t)
            }
        }
    }

    /** Corrections for a misspelled [word], closest first. */
    @ReactMethod
    fun spellSuggest(lang: String, word: String, promise: Promise) {
        worker.execute {
            try {
                promise.resolve(Arguments.fromList(SpellChecker.builtIn(lang).suggest(word)))
            } catch (t: Throwable) {
                promise.reject("DOCX_SPELL", t.message ?: t.toString(), t)
            }
        }
    }

    /** An EPUB's title, creators, date and publisher, for citing it. */
    @ReactMethod
    fun epubInfo(path: String, promise: Promise) {
        worker.execute {
            try {
                val info = DocxReader.epubInfo(File(path))
                promise.resolve(Arguments.createMap().apply {
                    putString("title", info.title)
                    putArray("creators", Arguments.fromList(info.creators))
                    putString("date", info.date)
                    putString("publisher", info.publisher)
                })
            } catch (t: Throwable) {
                promise.reject("DOCX_EPUB", t.message ?: t.toString(), t)
            }
        }
    }

    /** A small text file's contents (at most 16 KB): a key or setting the user keeps in a .txt. */
    @ReactMethod
    fun readText(path: String, promise: Promise) {
        worker.execute {
            try {
                val f = File(path)
                check(f.isFile) { "No such file: ${f.name}" }
                check(f.length() <= 16 * 1024) { "${f.name} is too large for a text snippet" }
                promise.resolve(f.readText(Charsets.UTF_8))
            } catch (t: Throwable) {
                promise.reject("DOCX_READ_TEXT", t.message ?: t.toString(), t)
            }
        }
    }

    /** Size and last-modified time of [path]: its fingerprint, to notice changes made elsewhere. */
    @ReactMethod
    fun fileStamp(path: String, promise: Promise) {
        worker.execute {
            val f = File(path)
            promise.resolve(if (f.isFile) "${f.length()}:${f.lastModified()}" else null)
        }
    }

    /**
     * A private copy of the document as it was opened. Edits are always applied to this
     * copy, so saving over the original any number of times stays correct.
     */
    @ReactMethod
    fun snapshot(path: String, key: String, promise: Promise) {
        worker.execute {
            try {
                val dest = File(home("snapshots"), safeName(key) + ".docx")
                File(path).inputStream().use { input -> FileOutputStream(dest).use { input.copyTo(it) } }
                promise.resolve(dest.path)
            } catch (t: Throwable) {
                appendLog("snapshot FAILED $path: $t")
                promise.reject("DOCX_SNAPSHOT", t.message ?: t.toString(), t)
            }
        }
    }

    /** Copies the current [path] into its backups before it is overwritten; keeps the newest five. */
    @ReactMethod
    fun backup(path: String, key: String, promise: Promise) {
        worker.execute {
            try {
                val dir = home("backups/" + safeName(key))
                val dest = File(dir, "${System.currentTimeMillis()}.docx")
                File(path).inputStream().use { input -> FileOutputStream(dest).use { input.copyTo(it) } }
                dir.listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedByDescending { it.name }.drop(5).forEach { it.delete() }
                promise.resolve(dest.path)
            } catch (t: Throwable) {
                appendLog("backup FAILED $path: $t")
                promise.reject("DOCX_BACKUP", t.message ?: t.toString(), t)
            }
        }
    }

    /** The document's backups, newest first: [{path, time, bytes}]. */
    @ReactMethod
    fun backups(key: String, promise: Promise) {
        worker.execute {
            val list = Arguments.createArray()
            home("backups/" + safeName(key)).listFiles { f -> f.name.endsWith(".docx") }.orEmpty().sortedByDescending { it.name }.forEach { f ->
                list.pushMap(Arguments.createMap().apply {
                    putString("path", f.path)
                    putDouble("time", f.nameWithoutExtension.toDoubleOrNull() ?: f.lastModified().toDouble())
                    putDouble("bytes", f.length().toDouble())
                })
            }
            promise.resolve(list)
        }
    }

    /** Where "Save a copy" of [path] goes: <name>-edited.docx beside it, never over a file. */
    @ReactMethod
    fun copyName(path: String, promise: Promise) {
        worker.execute { promise.resolve(DocxEditor.editedCopyName(File(path)).path) }
    }

    /** Writes [from] (a backup) over [dest] (the document). */
    @ReactMethod
    fun copyOver(from: String, dest: String, promise: Promise) {
        worker.execute {
            try {
                File(from).inputStream().use { input -> FileOutputStream(File(dest)).use { input.copyTo(it) } }
                appendLog("restored $from over $dest")
                promise.resolve(true)
            } catch (t: Throwable) {
                promise.reject("DOCX_COPY", t.message ?: t.toString(), t)
            }
        }
    }

    // ---------------------------------------------------------------- fonts

    /** Fonts added by hand, remembered across sessions: family \t style \t path. */
    private val fontList by lazy { File(reactContext.filesDir, "sn-docx-fonts.tsv") }
    private val loaded = HashMap<String, MutableSet<Int>>() // family → styles registered
    private var restored = false

    private val faces = HashMap<String, HashMap<Int, Typeface>>() // family → style → file-loaded face

    /** Stand-ins (as domain/fonts.ts): free fonts with the same widths as common Word fonts. */
    private val STAND_INS = mapOf(
        "calibri" to "Carlito",
        "cambria" to "Caladea",
        "arial" to "Liberation Sans",
        "helvetica" to "Liberation Sans",
        "times new roman" to "Liberation Serif",
        "times" to "Liberation Serif",
        "courier new" to "Liberation Mono",
        "courier" to "Liberation Mono",
        "georgia" to "Gelasio",
    )

    private fun register(family: String, style: Int, file: File): Boolean {
        val tf = runCatching { Typeface.createFromFile(file) }.getOrNull() ?: return false
        faces.getOrPut(family) { HashMap() }[style] = tf
        ReactFontManager.getInstance().setTypeface(family, style, tf)
        loaded.getOrPut(family) { HashSet() }.add(style)
        fillMissing(family)
        return true
    }

    /**
     * Styles with no file of their own (say, no bold file) get a synthesized face of the
     * regular one; otherwise the screen would draw them in the system font instead.
     */
    private fun fillMissing(family: String) {
        val have = faces[family] ?: return
        val base = have[Typeface.NORMAL] ?: have.values.first()
        for (style in listOf(Typeface.NORMAL, Typeface.BOLD, Typeface.ITALIC, Typeface.BOLD_ITALIC)) {
            if (style in have) continue
            // A real bold/italic from the font's metric-compatible stand-in beats a synthesized
            // one: the screen does not draw synthesized bold (CT: Times New Roman headings,
            // with only its regular file on the device, showed not bold).
            val borrowed = STAND_INS[family.lowercase()]?.let { FontFiles.find(it, style) }
                ?.let { runCatching { Typeface.createFromFile(it) }.getOrNull() }
            if (borrowed != null) appendLog("font $family: style $style borrowed from ${STAND_INS[family.lowercase()]}")
            ReactFontManager.getInstance().setTypeface(family, style, borrowed ?: Typeface.create(base, style))
        }
    }

    private fun restore() {
        if (restored) return
        restored = true
        runCatching {
            if (!fontList.isFile) return
            for (line in fontList.readLines()) {
                val (family, style, path) = line.split('\t').takeIf { it.size == 3 } ?: continue
                val f = File(path)
                if (f.isFile) register(family, style.toIntOrNull() ?: 0, f)
            }
        }
    }

    /**
     * Makes the given families (the fonts a document uses) available to the screen where a
     * file can be found in MyStyle/Fonts, plus every font added by hand. Resolves the
     * families that are available now.
     */
    @ReactMethod
    fun fonts(families: ReadableArray, promise: Promise) {
        worker.execute {
            restore()
            // Fonts may have been copied to the device since the last look.
            FontFiles.reindex()
            for (i in 0 until families.size()) {
                val family = families.getString(i) ?: continue
                // Style by style: a family loaded with only some styles (a regular added by hand,
                // say) still gets the rest from the font folder. Skipping the whole family once
                // anything was loaded left Times New Roman without its bold file (CT: headings
                // not bold on the device, though the bold file was there).
                for (style in 0..3) {
                    if (loaded[family]?.contains(style) == true) continue
                    FontFiles.find(family, style)?.let { register(family, style, it) }
                }
            }
            appendLog("fonts available: ${loaded.keys.sorted().joinToString { f -> "$f ${loaded[f]!!.sorted().joinToString("") { st -> "RBIZ"[st].toString() }}" }} (R regular, B bold, I italic, Z bold italic: files found)")
            promise.resolve(Arguments.fromList(loaded.keys.sorted()))
        }
    }

    /** A font file the user picked: named from its own name table, registered and remembered. */
    @ReactMethod
    fun addFont(path: String, promise: Promise) {
        worker.execute {
            restore()
            val file = File(path)
            val info = FontFiles.info(file)
            if (info == null) {
                promise.reject("DOCX_FONT", "Not a font file this device can read: ${file.name}")
                return@execute
            }
            val style = FontFiles.styleIndex(info.bold, info.italic)
            if (!register(info.family, style, file)) {
                promise.reject("DOCX_FONT", "Android could not load ${file.name}")
                return@execute
            }
            runCatching { fontList.appendText("${info.family}\t$style\t${file.path}\n") }
            appendLog("font added: ${info.family} style $style from ${file.path}")
            promise.resolve(Arguments.createMap().apply {
                putString("family", info.family)
                putInt("style", style)
            })
        }
    }

    private fun report(r: DocxReader.Report): WritableMap = Arguments.createMap().apply {
        putInt("paragraphs", r.paragraphs)
        putInt("tables", r.tables)
        putInt("images", r.images)
        putInt("trackedChanges", r.trackedChanges)
        putInt("comments", r.comments)
        putInt("fields", r.fields)
        putInt("contentControls", r.contentControls)
    }

    private fun notePieces(pieces: List<DocxReader.NotePiece>) = Arguments.createArray().apply {
        for (p in pieces) pushMap(Arguments.createMap().apply {
            putString("t", p.text)
            if (p.italic) putBoolean("i", true)
            if (p.bold) putBoolean("b", true)
        })
    }

    /** Relationship id → extracted file, for the open document's pictures. */
    @Volatile
    private var pictures: Map<String, String> = emptyMap()

    private fun runs(list: List<DocxReader.Run>) = Arguments.createArray().apply {
        for (r in list) pushMap(Arguments.createMap().apply {
            putString("t", r.text)
            // Bold, italic, font and size as the run sets them itself; the style's part comes
            // with the paragraph (sb/si/bf/bs), so a style change shows on screen at once.
            r.ownBold?.let { putBoolean("b", it) }
            r.ownItalic?.let { putBoolean("i", it) }
            if (r.underline) putBoolean("u", true)
            if (r.strike) putBoolean("s", true)
            if (r.highlight) putBoolean("h", true)
            if (r.link) putBoolean("l", true)
            if (r.superscript) putBoolean("sup", true)
            if (r.subscript) putBoolean("sub", true)
            r.highlightColor?.let { putString("hc", it) }
            r.obj?.let { putString("obj", it) }
            if (r.locked) putBoolean("k", true)
            r.ownFont?.let { putString("f", it) }
            r.ownSize?.let { putInt("sz", it) }
            r.rev?.let { putString("rv", it) }
            r.ink?.let { putString("ink", it) }
            if (r.pageBreak) putBoolean("pg", true)
            r.imageRel?.let { pictures[it] }?.let { putString("src", it) }
            r.noteId?.let { putString("fn", it) }
            if (r.cx > 0 && r.cy > 0) {
                putDouble("cx", r.cx.toDouble())
                putDouble("cy", r.cy.toDouble())
            }
        })
    }

    private fun block(b: DocxReader.Block): WritableMap = Arguments.createMap().apply {
        when (b) {
            is DocxReader.Paragraph -> {
                putString("type", "p")
                putInt("index", b.index)
                putString("style", b.styleId)
                putString("kind", b.kind)
                putInt("level", b.level)
                putString("align", b.align)
                putInt("indent", b.indentTwips)
                b.listLabel?.let { putString("list", it) }
                if (b.sectionBreak) putBoolean("sect", true)
                if (b.quote) putBoolean("quote", true)
                if (b.ownAlign) putBoolean("ja", true)
                if (b.ownIndent) putBoolean("ji", true)
                b.numId?.let { id -> putMap("num", Arguments.createMap().apply { putInt("id", id); putInt("lvl", b.ilvl) }) }
                b.para.let { f ->
                    f.before?.let { putInt("before", it) }
                    f.after?.let { putInt("after", it) }
                    f.line?.let { putInt("line", it) }
                    f.lineRule?.let { putString("lineRule", it) }
                    f.first?.let { putInt("first", it) }
                    if (f.pageBreakBefore == true) putBoolean("pb", true)
                }
                b.baseFont?.let { putString("bf", it) }
                b.baseSize?.let { putInt("bs", it) }
                if (b.baseBold) putBoolean("sb", true)
                if (b.baseItalic) putBoolean("si", true)
                putArray("runs", runs(b.runs))
                if (b.revisions.isNotEmpty()) putArray("revs", Arguments.createArray().apply {
                    for (v in b.revisions) pushMap(Arguments.createMap().apply {
                        putString("id", v.id)
                        putString("kind", v.kind)
                        putString("author", v.author)
                        putString("date", v.date)
                        if (v.kind == "del") {
                            putInt("at", v.at)
                            putArray("runs", runs(v.runs))
                        }
                        if (v.move) putBoolean("move", true)
                    })
                })
                if (b.marks.isNotEmpty()) putArray("marks", Arguments.createArray().apply {
                    for (k in b.marks) pushMap(Arguments.createMap().apply {
                        putString("id", k.id)
                        putString("kind", k.kind)
                        putInt("at", k.at)
                    })
                })
            }
            is DocxReader.Table -> {
                putString("type", "table")
                putInt("rows", b.rows)
                putInt("cols", b.cols)
                putString("preview", b.preview)
                putInt("t", b.index)
                putArray("widths", Arguments.createArray().apply { b.widths.forEach { pushInt(it) } })
                if (b.widthFrac < 0.995) putDouble("wf", b.widthFrac)
                if (b.align != "left") putString("ta", b.align)
                putArray("grid", Arguments.createArray().apply {
                    for (row in b.grid) pushArray(Arguments.createArray().apply {
                        for (c in row) pushMap(Arguments.createMap().apply {
                            if (c.span > 1) putInt("s", c.span)
                            if (c.merged) putBoolean("m", true)
                            if (c.nested) putBoolean("n", true)
                            if (c.rich) putBoolean("r", true)
                            putArray("p", notePieces(c.pieces))
                        })
                    })
                })
            }
            is DocxReader.Protected -> {
                putString("type", "protected")
                putString("what", b.what)
                putString("preview", b.preview)
            }
        }
    }
}
