package com.sndocx

import android.graphics.Color
import android.view.View
import com.facebook.react.uimanager.SimpleViewManager
import com.facebook.react.uimanager.ThemedReactContext

/**
 * Host view for the firmware pen engine (handwritten notes), rendered from JSX as
 * `<DocxInkSurface />` so React lays it out — the engine sizes its writing bitmap from the
 * host's measured size. A plain View (the engine adds to its overlay). Give it a FIXED size
 * and never move it while the engine is live. Ported from sn-flashcards' InkSurface; see
 * references/firmware-pen-engine-einkpwinterface.md.
 */
class DocxInkViewManager : SimpleViewManager<View>() {

    override fun getName() = "DocxInkSurface"

    override fun createViewInstance(context: ThemedReactContext): View {
        val v = View(context)
        v.setBackgroundColor(Color.WHITE)
        current = v
        return v
    }

    /** An engine left enabled against a detached view wedges PluginHost: release first. */
    override fun onDropViewInstance(view: View) {
        if (current === view) {
            DocxInkModule.release("ink host unmounted")
            current = null
        }
        super.onDropViewInstance(view)
    }

    companion object {
        @Volatile
        var current: View? = null
            private set
    }
}
