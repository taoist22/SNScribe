package com.sndocx

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.view.KeyEvent
import android.view.View
import com.facebook.react.bridge.Arguments
import com.facebook.react.bridge.Promise
import com.facebook.react.bridge.ReactApplicationContext
import com.facebook.react.bridge.ReactContextBaseJavaModule
import com.facebook.react.bridge.ReactMethod
import com.facebook.react.bridge.UiThreadUtil
import com.facebook.react.modules.core.DeviceEventManagerModule
import com.facebook.react.uimanager.UIManagerHelper

/**
 * Keyboard keys the editor handles itself, and the clipboard.
 *
 * A plugin has no Activity, so no onKeyDown; instead an OnKeyListener sits on the editor's
 * invisible text field (resolved by react tag). Hardware keys reach the focused view — the
 * probe (sn-docx-probe, spike S2) saw arrows, Esc and Ctrl combinations arrive this way,
 * while typed letters and Backspace come through the input method as text.
 *
 * The listener consumes only the keys below, so the text field never acts on them, and
 * sends each press to JS as "DocxKey" {key, shift, text?}. Ctrl and Cmd (Meta) both count,
 * so a keyboard set up for a Mac works too.
 */
class DocxKeysModule(private val reactContext: ReactApplicationContext) : ReactContextBaseJavaModule(reactContext) {

    override fun getName() = "DocxKeys"

    private val names = mapOf(
        KeyEvent.KEYCODE_FORWARD_DEL to "DEL_FWD",
        KeyEvent.KEYCODE_DPAD_LEFT to "LEFT",
        KeyEvent.KEYCODE_DPAD_RIGHT to "RIGHT",
        KeyEvent.KEYCODE_DPAD_UP to "UP",
        KeyEvent.KEYCODE_DPAD_DOWN to "DOWN",
        KeyEvent.KEYCODE_MOVE_HOME to "HOME",
        KeyEvent.KEYCODE_MOVE_END to "END",
        // Tab would move focus off the text field (and drop the keyboard); it types a tab instead.
        KeyEvent.KEYCODE_TAB to "TAB",
    )

    private val shortcuts = mapOf(
        KeyEvent.KEYCODE_C to "C",
        KeyEvent.KEYCODE_X to "X",
        KeyEvent.KEYCODE_V to "V",
        KeyEvent.KEYCODE_Z to "Z",
        KeyEvent.KEYCODE_Y to "Y",
        KeyEvent.KEYCODE_B to "B",
        KeyEvent.KEYCODE_I to "I",
        KeyEvent.KEYCODE_U to "U",
    )

    private val listener = View.OnKeyListener { _, keyCode, event ->
        val mod = event.isCtrlPressed || event.isMetaPressed
        val key = names[keyCode] ?: (if (mod) shortcuts[keyCode] else null) ?: return@OnKeyListener false
        var paste: String? = null
        if (key == "V") {
            // Paste ourselves when the clipboard can be read; otherwise let the field paste.
            paste = clipboardText() ?: return@OnKeyListener false
        }
        if (event.action == KeyEvent.ACTION_DOWN) {
            val map = Arguments.createMap().apply {
                putString("key", key)
                putBoolean("shift", event.isShiftPressed)
                paste?.let { putString("text", it) }
            }
            runCatching {
                reactContext.getJSModule(DeviceEventManagerModule.RCTDeviceEventEmitter::class.java).emit("DocxKey", map)
            }
        }
        true
    }

    /** Listens on the text field with this react tag. Resolves what it attached to. */
    @ReactMethod
    fun attach(tag: Double, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val reactTag = tag.toInt()
            val view = runCatching {
                UIManagerHelper.getUIManagerForReactTag(reactContext, reactTag)?.resolveView(reactTag)
            }.getOrNull()
            if (view == null) {
                promise.resolve("no view for tag $reactTag")
                return@runOnUiThread
            }
            view.setOnKeyListener(listener)
            promise.resolve("keys on ${view.javaClass.simpleName}")
        }
    }

    @ReactMethod
    fun copy(text: String, promise: Promise) {
        UiThreadUtil.runOnUiThread {
            val ok = runCatching {
                clipboard()?.setPrimaryClip(ClipData.newPlainText("DOCX", text)) != null
            }.getOrDefault(false)
            promise.resolve(ok)
        }
    }

    private fun clipboard() = reactContext.getSystemService(Context.CLIPBOARD_SERVICE) as? ClipboardManager

    private fun clipboardText(): String? = runCatching {
        clipboard()?.primaryClip?.takeIf { it.itemCount > 0 }?.getItemAt(0)?.coerceToText(reactContext)?.toString()
    }.getOrNull()

    // NativeEventEmitter bookkeeping (events go out through RCTDeviceEventEmitter regardless).
    @ReactMethod
    fun addListener(eventName: String) {}

    @ReactMethod
    fun removeListeners(count: Double) {}
}
