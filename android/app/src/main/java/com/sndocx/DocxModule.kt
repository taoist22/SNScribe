package com.sndocx

import com.facebook.react.bridge.Arguments
import com.facebook.react.bridge.Promise
import com.facebook.react.bridge.ReactApplicationContext
import com.facebook.react.bridge.ReactContextBaseJavaModule
import com.facebook.react.bridge.ReactMethod
import com.facebook.react.bridge.WritableMap
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
class DocxModule(reactContext: ReactApplicationContext) : ReactContextBaseJavaModule(reactContext) {

    override fun getName() = "Docx"

    private val worker = Executors.newSingleThreadExecutor()
    private val timer = Executors.newSingleThreadScheduledExecutor()

    companion object {
        /** Bumped with each native change; first log line, to spot stale installs. */
        const val NATIVE_BUILD = 1
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
                val ms = System.currentTimeMillis() - t0
                appendLog("open ${file.name}: ${file.length()} B, ${result.blocks.size} blocks, $ms ms, ${result.report}")
                promise.resolve(Arguments.createMap().apply {
                    putString("path", file.path)
                    putString("name", file.name)
                    putDouble("bytes", file.length().toDouble())
                    putDouble("ms", ms.toDouble())
                    putMap("report", report(result.report))
                    putArray("blocks", Arguments.createArray().apply { result.blocks.forEach { pushMap(block(it)) } })
                })
            } catch (t: Throwable) {
                appendLog("open FAILED $path: $t")
                promise.reject("DOCX_OPEN_FAILED", t.message ?: t.toString(), t)
            }
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
                putArray("runs", Arguments.createArray().apply {
                    for (r in b.runs) pushMap(Arguments.createMap().apply {
                        putString("t", r.text)
                        if (r.bold) putBoolean("b", true)
                        if (r.italic) putBoolean("i", true)
                        if (r.underline) putBoolean("u", true)
                        if (r.strike) putBoolean("s", true)
                        if (r.highlight) putBoolean("h", true)
                        if (r.link) putBoolean("l", true)
                        if (r.superscript) putBoolean("sup", true)
                        r.obj?.let { putString("obj", it) }
                    })
                })
            }
            is DocxReader.Table -> {
                putString("type", "table")
                putInt("rows", b.rows)
                putInt("cols", b.cols)
                putString("preview", b.preview)
            }
            is DocxReader.Protected -> {
                putString("type", "protected")
                putString("what", b.what)
                putString("preview", b.preview)
            }
        }
    }
}
