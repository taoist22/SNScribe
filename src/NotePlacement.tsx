// Add to note: the selected passage goes into the note SNScribe was opened from, as a
// Supernote text box (searchable, editable), where the pen taps. SNScribe steps aside while
// placing: this view is transparent, so the note's current page shows, and it owns every
// touch, so the note never inks. Pattern and baseline calibration from sn-datetime
// (PlacementOverlay, tapPlacement); insertText coordinates are device pixels.

import React, {useEffect, useRef} from 'react';
import {BackHandler, PixelRatio, Pressable, StyleSheet, Text, View} from 'react-native';
import {PluginCommAPI, PluginNoteAPI} from 'sn-plugin-lib';

type ApiRes<T> = {success?: boolean; result?: T; error?: {message?: string}};

/** Text size of the box (Supernote's text font size). */
const FONT_SIZE = 32;
/** Distance from the box top to the baseline, as a multiple of the font size (sn-datetime, measured). */
const BASELINE_OFFSET_RATIO = 2.05;

/** The note the plugin was opened from (a .note file), or null (a PDF or EPUB). */
export async function openNotePath(): Promise<string | null> {
  try {
    const res = (await PluginCommAPI.getCurrentFilePath()) as ApiRes<string>;
    const path = res?.success ? res.result : undefined;
    return typeof path === 'string' && /\.note$/i.test(path) ? path : null;
  } catch {
    return null;
  }
}

async function pageSize(): Promise<{width: number; height: number}> {
  try {
    const res = (await PluginCommAPI.getPageDisplaySize()) as ApiRes<{width: number; height: number}>;
    if (res?.success && res.result?.width && res.result.height) {
      return res.result;
    }
  } catch {
    // Older firmware: the Manta's page.
  }
  return {width: 1920, height: 2560};
}

/** Puts `text` into the open note with its first line's baseline on the tapped point. */
export async function insertIntoNote(text: string, point: {x: number; y: number}): Promise<void> {
  const page = await pageSize();
  const margin = Math.round(page.width * 0.06);
  const left = Math.max(margin, Math.min(point.x, page.width - margin - 400));
  const right = page.width - margin;
  const top = Math.max(0, Math.round(point.y - FONT_SIZE * BASELINE_OFFSET_RATIO));
  const res = (await PluginNoteAPI.insertText({
    textContentFull: text,
    textRect: {left: Math.round(left), top, right: Math.round(right), bottom: Math.min(page.height, top + FONT_SIZE * 3)},
    fontSize: FONT_SIZE,
    textBold: 0,
    textItalics: 0,
    textAlign: 0,
    textEditable: 1,
    showLassoAfterInsert: false,
  })) as ApiRes<boolean>;
  if (!res?.success) {
    throw new Error(res?.error?.message ?? 'The note did not take the text.');
  }
}

/** Tap to place; Cancel, Back or 30 seconds without a tap cancel. */
export function NotePlacement({onPlace, onCancel}: {onPlace: (point: {x: number; y: number}) => void; onCancel: () => void}) {
  const surface = useRef<View>(null);
  const origin = useRef({x: 0, y: 0});
  const down = useRef<{x: number; y: number; time: number} | null>(null);
  const settled = useRef(false);
  const cancel = () => {
    if (!settled.current) {
      settled.current = true;
      onCancel();
    }
  };
  useEffect(() => {
    const timer = setTimeout(cancel, 30000);
    const back = BackHandler.addEventListener('hardwareBackPress', () => {
      cancel();
      return true;
    });
    return () => {
      clearTimeout(timer);
      back.remove();
    };
    // Mounted once per placement.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <View style={styles.root}>
      <View
        ref={surface}
        collapsable={false}
        style={StyleSheet.absoluteFill}
        onLayout={() => surface.current?.measureInWindow((x, y) => (origin.current = {x, y}))}
        onStartShouldSetResponder={() => true}
        onMoveShouldSetResponder={() => true}
        onResponderTerminationRequest={() => false}
        onResponderGrant={({nativeEvent: e}) => {
          down.current = e.touches.length === 1 ? {x: e.locationX, y: e.locationY, time: e.timestamp} : null;
        }}
        onResponderMove={({nativeEvent: e}) => {
          if (e.touches.length !== 1 || (down.current && Math.hypot(e.locationX - down.current.x, e.locationY - down.current.y) > 20)) {
            down.current = null;
          }
        }}
        onResponderTerminate={() => {
          down.current = null;
        }}
        onResponderRelease={({nativeEvent: e}) => {
          const start = down.current;
          down.current = null;
          if (settled.current || !start || e.timestamp - start.time > 700 || Math.hypot(e.locationX - start.x, e.locationY - start.y) > 20) {
            return;
          }
          // Touches are in dp; the note wants device pixels.
          const scale = PixelRatio.get();
          settled.current = true;
          onPlace({x: (origin.current.x + e.locationX) * scale, y: (origin.current.y + e.locationY) * scale});
        }}
      />
      <View style={styles.hint}>
        <Text allowFontScaling={false} style={styles.text}>
          {'Tap where the passage goes in your note'}
        </Text>
        <Pressable onPress={cancel} style={styles.cancel}>
          <Text allowFontScaling={false} style={styles.text}>
            {'Cancel'}
          </Text>
        </Pressable>
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {flex: 1, backgroundColor: 'transparent'},
  hint: {position: 'absolute', top: 10, alignSelf: 'center', flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', borderWidth: 1, borderColor: '#000', paddingLeft: 12},
  text: {color: '#000', fontSize: 17},
  cancel: {padding: 16, marginLeft: 12, borderLeftWidth: 1, borderColor: '#000'},
});
