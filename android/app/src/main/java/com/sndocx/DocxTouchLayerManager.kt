package com.sndocx

import android.content.Context
import android.view.MotionEvent
import com.facebook.react.uimanager.ThemedReactContext
import com.facebook.react.views.view.ReactViewGroup
import com.facebook.react.views.view.ReactViewManager

/**
 * The page's touch layer, `<DocxTouchLayer />`: an ordinary React Native view that also
 * notes whether the pen or a finger made the current touch. Android reports the tool on
 * every MotionEvent; React Native's touch events leave it out. JS asks afterwards
 * (DocxText.gestureTool) and lets the pen select while a finger only turns pages.
 *
 * On Supernote the pen reports TOOL_TYPE_STYLUS (sn-lastnote, sn-flashcards' pen engine).
 */
class DocxTouchLayerManager : ReactViewManager() {

    override fun getName() = "DocxTouchLayer"

    override fun createViewInstance(context: ThemedReactContext): ReactViewGroup = TouchLayer(context)

    class TouchLayer(context: Context) : ReactViewGroup(context) {
        override fun dispatchTouchEvent(ev: MotionEvent): Boolean {
            when (ev.actionMasked) {
                MotionEvent.ACTION_DOWN -> {
                    gesturePen = isPen(ev, 0)
                    lastTool = ev.getToolType(0)
                }
                // A pen that joins a resting hand makes it a pen gesture.
                MotionEvent.ACTION_POINTER_DOWN -> if (isPen(ev, ev.actionIndex)) gesturePen = true
            }
            return super.dispatchTouchEvent(ev)
        }
    }

    companion object {
        /** Whether the touch that started last is (or came to include) the pen. */
        @Volatile
        var gesturePen: Boolean = true

        /** Android's tool type of the last touch-down, for the log. */
        @Volatile
        var lastTool: Int = MotionEvent.TOOL_TYPE_UNKNOWN

        /** Anything but a finger (or palm, Android 13's tool type 5) counts as the pen, so an unknown tool still selects. */
        private fun isPen(ev: MotionEvent, i: Int): Boolean {
            val tool = ev.getToolType(i)
            return tool != MotionEvent.TOOL_TYPE_FINGER && tool != TOOL_TYPE_PALM
        }

        private const val TOOL_TYPE_PALM = 5
    }
}
