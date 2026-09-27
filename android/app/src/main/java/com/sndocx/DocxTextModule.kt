package com.sndocx

import android.widget.TextView
import com.facebook.react.bridge.Arguments
import com.facebook.react.bridge.Promise
import com.facebook.react.bridge.ReactApplicationContext
import com.facebook.react.bridge.ReactContextBaseJavaModule
import com.facebook.react.bridge.ReactMethod
import com.facebook.react.bridge.UiThreadUtil
import com.facebook.react.uimanager.UIManagerHelper

/**
 * Which character of a paragraph is under the pen? (Proven in sn-docx-probe spike S1b.)
 *
 * A paragraph is drawn as one RN Text (fast — spike S1), so there are no per-word boxes.
 * Instead this asks the paragraph's own TextView for the android.text.Layout it drew with,
 * and maps a point to a character offset with it. Offsets are UTF-16 indexes into the
 * paragraph's text, the same indexes DocxFileModule edits by.
 */
class DocxTextModule(private val reactContext: ReactApplicationContext) :
    ReactContextBaseJavaModule(reactContext) {

    override fun getName() = "DocxText"

    /**
     * x, y: pixels relative to the Text view's top-left. Resolves
     * {offset (nearest gap), char (character under the pen), line, lineCount, length, viewClass}
     * or {error, viewClass}; never rejects.
     */
    @ReactMethod
    fun offsetAt(tag: Double, x: Double, y: Double, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val result = Arguments.createMap()
            try {
                val reactTag = tag.toInt()
                val view = UIManagerHelper.getUIManagerForReactTag(reactContext, reactTag)?.resolveView(reactTag)
                result.putString("viewClass", view?.javaClass?.name ?: "null")
                val layout = (view as? TextView)?.layout
                if (view !is TextView) {
                    result.putString("error", "not a TextView")
                } else if (layout == null) {
                    result.putString("error", "TextView has no layout yet")
                } else {
                    val lx = x.toFloat() - view.totalPaddingLeft
                    val ly = y.toInt() - view.totalPaddingTop
                    val line = layout.getLineForVertical(ly)
                    val offset = layout.getOffsetForHorizontal(line, lx)
                    // offset is the nearest gap between characters; the character under the
                    // pen is the one before it when the gap lies right of the pen.
                    val lineStart = layout.getLineStart(line)
                    val lineEnd = layout.getLineEnd(line)
                    var ch = if (offset > lineStart && layout.getPrimaryHorizontal(offset) > lx) offset - 1 else offset
                    ch = ch.coerceIn(lineStart, maxOf(lineStart, lineEnd - 1))
                    result.putInt("offset", offset)
                    result.putInt("char", ch)
                    result.putInt("line", line)
                    result.putInt("lineCount", layout.lineCount)
                    result.putInt("length", view.text?.length ?: 0)
                }
            } catch (t: Throwable) {
                result.putString("error", t.toString())
            }
            promise.resolve(result)
        }
    }
    /**
     * Where to draw a caret before character [offset]: {x, top, bottom} in px relative to the
     * Text view (padding included), or {error}. Never rejects.
     */
    @ReactMethod
    fun caretRect(tag: Double, offset: Double, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val result = Arguments.createMap()
            try {
                val reactTag = tag.toInt()
                val view = UIManagerHelper.getUIManagerForReactTag(reactContext, reactTag)?.resolveView(reactTag)
                val layout = (view as? TextView)?.layout
                if (view !is TextView || layout == null) {
                    result.putString("error", "no text layout")
                } else {
                    val o = offset.toInt().coerceIn(0, layout.text.length)
                    val line = layout.getLineForOffset(o)
                    result.putDouble("x", (layout.getPrimaryHorizontal(o) + view.totalPaddingLeft).toDouble())
                    result.putDouble("top", (layout.getLineTop(line) + view.totalPaddingTop).toDouble())
                    result.putDouble("bottom", (layout.getLineBottom(line) + view.totalPaddingTop).toDouble())
                }
            } catch (t: Throwable) {
                result.putString("error", t.toString())
            }
            promise.resolve(result)
        }
    }
    /**
     * The caret's offset one line up (dir -1) or down (+1), keeping its x — or
     * {outside: true} when that line is past the paragraph. Never rejects.
     */
    @ReactMethod
    fun lineMove(tag: Double, offset: Double, dir: Double, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val result = Arguments.createMap()
            try {
                val reactTag = tag.toInt()
                val view = UIManagerHelper.getUIManagerForReactTag(reactContext, reactTag)?.resolveView(reactTag)
                val layout = (view as? TextView)?.layout
                if (layout == null) {
                    result.putString("error", "no text layout")
                } else {
                    val o = offset.toInt().coerceIn(0, layout.text.length)
                    val target = layout.getLineForOffset(o) + dir.toInt()
                    if (target < 0 || target >= layout.lineCount) {
                        result.putBoolean("outside", true)
                    } else {
                        result.putInt("offset", layout.getOffsetForHorizontal(target, layout.getPrimaryHorizontal(o)))
                    }
                }
            } catch (t: Throwable) {
                result.putString("error", t.toString())
            }
            promise.resolve(result)
        }
    }
}
