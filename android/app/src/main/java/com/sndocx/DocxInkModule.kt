package com.sndocx

import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Path
import android.view.View
import com.facebook.react.bridge.Arguments
import com.facebook.react.bridge.Promise
import com.facebook.react.bridge.ReactApplicationContext
import com.facebook.react.bridge.ReactContextBaseJavaModule
import com.facebook.react.bridge.ReactMethod
import com.facebook.react.bridge.UiThreadUtil
import java.io.File
import java.io.FileOutputStream

/**
 * Handwritten notes: the firmware's pen engine on a plugin surface (as in sn-flashcards and
 * the sn-docx-probe T1 test), and the strokes read back and drawn into a PNG.
 *
 * Undocumented vendor reflection, so every entry point fails soft; [isAvailable] gates the
 * feature. Every exit path must [release] the engine, or PluginHost wedges.
 */
class DocxInkModule(private val reactContext: ReactApplicationContext) : ReactContextBaseJavaModule(reactContext) {

    override fun getName() = "DocxInk"

    override fun invalidate() {
        release("react instance torn down")
        super.invalidate()
    }

    companion object {
        private const val PENTYPE_BALLPEN = 7
        private const val DRAWOBJ_FREEHAND = 0
        private const val PENCOLOR_BLACK = -16777216
        private const val DISPATH_TOUCH_HOST_LATER = 2
        private const val PEN_WIDTH = 8.0f

        @Volatile private var engine: Any? = null

        /** Every exit path comes through here. Never mints an engine just to stop one. */
        fun release(reason: String): String {
            val e = engine ?: return "nothing active"
            val r = runCatching { e.javaClass.getMethod("setPWEnabled", java.lang.Boolean.TYPE).invoke(e, false) }
            engine = null
            android.util.Log.i("DocxInk", "released: $reason")
            return if (r.isSuccess) "released" else "release threw: ${r.exceptionOrNull()?.cause}"
        }

        /**
         * Draws strokes (x, y lists, in the engine's pixels) into a PNG cropped to the ink, as
         * thick as the pen wrote them and smoothed between samples (thin, straight-segment
         * lines broke up — "grainy" — once Word shrank the picture into the margin), in
         * [color] on a transparent background (so a note never hides text it overlaps in
         * Word). Returns (width, height), or null when there is no ink.
         */
        fun renderPng(strokes: List<IntArray>, dest: File, color: Int = Color.BLACK, strokeWidth: Float = PEN_WIDTH, maxSide: Int = 1000): Pair<Int, Int>? {
            val points = strokes.filter { it.size >= 2 }
            if (points.isEmpty()) return null
            var minX = Int.MAX_VALUE; var minY = Int.MAX_VALUE; var maxX = Int.MIN_VALUE; var maxY = Int.MIN_VALUE
            for (s in points) for (i in s.indices step 2) {
                minX = minOf(minX, s[i]); maxX = maxOf(maxX, s[i])
                minY = minOf(minY, s[i + 1]); maxY = maxOf(maxY, s[i + 1])
            }
            val pad = (strokeWidth + 8).toInt()
            val w0 = maxX - minX + 2 * pad
            val h0 = maxY - minY + 2 * pad
            val k = minOf(1f, maxSide.toFloat() / maxOf(w0, h0))
            val w = maxOf(1, (w0 * k).toInt())
            val h = maxOf(1, (h0 * k).toInt())
            val bmp = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bmp)
            val ink = color
            val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                this.color = ink
                style = Paint.Style.STROKE
                this.strokeWidth = maxOf(3f, strokeWidth * k)
                strokeCap = Paint.Cap.ROUND
                strokeJoin = Paint.Join.ROUND
            }
            for (s in points) {
                fun x(i: Int) = (s[i] - minX + pad) * k
                fun y(i: Int) = (s[i + 1] - minY + pad) * k
                if (s.size == 2) {
                    canvas.drawPoint(x(0), y(0), paint)
                    continue
                }
                // Through the midpoints, with each sample as a control point: a smooth curve.
                val path = Path()
                path.moveTo(x(0), y(0))
                var i = 2
                while (i + 2 < s.size) {
                    path.quadTo(x(i), y(i), (x(i) + x(i + 2)) / 2, (y(i) + y(i + 2)) / 2)
                    i += 2
                }
                path.lineTo(x(s.size - 2), y(s.size - 2))
                canvas.drawPath(path, paint)
            }
            dest.parentFile?.mkdirs()
            FileOutputStream(dest).use { bmp.compress(Bitmap.CompressFormat.PNG, 100, it) }
            bmp.recycle()
            return w to h
        }
    }

    @ReactMethod
    fun isAvailable(promise: Promise) {
        val ok = runCatching {
            View::class.java.getMethod("getPWInterFace")
            Class.forName("android.view.EinkPWInterface")
            true
        }.getOrDefault(false)
        promise.resolve(ok)
    }

    /** Bind to the mounted surface and take stylus input (finger off = palm rejection). */
    @ReactMethod
    fun activate(promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val host = DocxInkViewManager.current
            if (host == null) { promise.resolve("no-surface"); return@runOnUiThread }
            if (host.width == 0 || host.height == 0) { promise.resolve("not-laid-out"); return@runOnUiThread }
            val e = engine ?: runCatching { View::class.java.getMethod("getPWInterFace").invoke(host) }.getOrNull()
            if (e == null) { promise.resolve("no-engine"); return@runOnUiThread }
            engine = e
            runCatching { e.javaClass.getMethod("onWindowChanged").invoke(e) }
            runCatching { e.javaClass.getMethod("clearUnWriteRectList").invoke(e) }
            set(e, "setPenType", Integer.TYPE, PENTYPE_BALLPEN)
            set(e, "setDrawObjectType", Integer.TYPE, DRAWOBJ_FREEHAND)
            set(e, "setPenStdWidth", java.lang.Float.TYPE, PEN_WIDTH)
            set(e, "setPenColor", Integer.TYPE, PENCOLOR_BLACK)
            set(e, "setFingerWritable", java.lang.Boolean.TYPE, false)
            set(e, "disablePenEraseFloatFlower", java.lang.Boolean.TYPE, true)
            set(e, "enableTouchDispatch", Integer.TYPE, DISPATH_TOUCH_HOST_LATER)
            set(e, "disablePWInput", java.lang.Boolean.TYPE, false)
            val on = runCatching { e.javaClass.getMethod("setPWEnabled", java.lang.Boolean.TYPE).invoke(e, true) }
            // Clear after binding (not before), then force a repaint.
            runCatching { e.javaClass.getMethod("clearContentX", java.lang.Boolean.TYPE).invoke(e, true) }
            repaint(e, host)
            promise.resolve(if (on.isSuccess) "active" else "enable-failed")
        }
    }

    @ReactMethod
    fun clear(promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val e = engine
            val host = DocxInkViewManager.current
            if (e == null || host == null) { promise.resolve("inactive"); return@runOnUiThread }
            val r = runCatching { e.javaClass.getMethod("clearContentX", java.lang.Boolean.TYPE).invoke(e, true) }
            repaint(e, host)
            promise.resolve(if (r.isSuccess) "cleared" else "clear-failed")
        }
    }

    @ReactMethod
    fun deactivate(promise: Promise) {
        UiThreadUtil.runOnUiThread { promise.resolve(release("note pad closed")) }
    }

    /**
     * The ink written so far as a PNG at [path] in [color] ("#RRGGBB"; the pad itself can
     * only show black), cropped to the writing. Resolves
     * {path, width, height, strokes} or {empty: true}.
     */
    @ReactMethod
    fun save(path: String, color: String, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val e = engine
            if (e == null) { promise.reject("DOCX_INK_INACTIVE", "The pen pad is not active."); return@runOnUiThread }
            try {
                val steps = (e.javaClass.getMethod("getAvailableRepaintStep").invoke(e) as? Int) ?: 0
                val getStep = e.javaClass.getMethod("getStepPointArray", Integer.TYPE)
                val strokes = ArrayList<IntArray>()
                for (step in 0 until steps) {
                    val list = getStep.invoke(e, step) as? java.util.ArrayList<*> ?: continue
                    val xy = ArrayList<Int>()
                    for (pt in list) {
                        if (pt == null) continue
                        val action = intField(pt, "action") ?: continue
                        if (action and 0xFF00 == 0x7C00) continue // config records carry no position
                        if (intField(pt, "erasing")?.let { it != 0 } == true) continue
                        xy.add(intField(pt, "x") ?: continue)
                        xy.add(intField(pt, "y") ?: 0)
                    }
                    if (xy.isNotEmpty()) strokes.add(xy.toIntArray())
                }
                val dest = File(path)
                val ink = runCatching { Color.parseColor(color) }.getOrDefault(Color.BLACK)
                val size = renderPng(strokes, dest, ink)
                promise.resolve(Arguments.createMap().apply {
                    if (size == null) {
                        putBoolean("empty", true)
                    } else {
                        putString("path", dest.path)
                        putInt("width", size.first)
                        putInt("height", size.second)
                        putInt("strokes", strokes.size)
                    }
                })
            } catch (t: Throwable) {
                promise.reject("DOCX_INK_SAVE_FAILED", (t.cause ?: t).toString(), t)
            }
        }
    }

    /** Where a document's handwritten notes are kept (private storage); created if missing. */
    @ReactMethod
    fun notesDir(key: String, promise: Promise) {
        val safe = key.replace(Regex("[^A-Za-z0-9_-]"), "_").take(80)
        val dir = File(reactContext.filesDir, "sn-docx/notes/$safe").also { it.mkdirs() }
        promise.resolve(dir.path)
    }

    /** Removes a note picture from private storage (only there). */
    @ReactMethod
    fun remove(path: String, promise: Promise) {
        val f = File(path)
        val root = File(reactContext.filesDir, "sn-docx/notes").canonicalPath
        promise.resolve(f.canonicalPath.startsWith(root) && f.delete())
    }

    private fun repaint(e: Any, host: View) {
        runCatching {
            e.javaClass.getMethod("invalidateHost", android.graphics.Rect::class.java)
                .invoke(e, android.graphics.Rect(0, 0, host.width, host.height))
        }
        host.invalidate()
    }

    private fun set(engine: Any, name: String, type: Class<*>, value: Any) {
        runCatching { engine.javaClass.getMethod(name, type).invoke(engine, value) }
            .onFailure { android.util.Log.w("DocxInk", "$name failed: ${it.cause ?: it}") }
    }

    private fun intField(pt: Any, name: String): Int? = runCatching { pt.javaClass.getField(name).getInt(pt) }.getOrNull()
}
