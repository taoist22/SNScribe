import React, {useEffect, useMemo, useRef, useState} from 'react';
import {StyleSheet, View, type LayoutChangeEvent} from 'react-native';
import {BlockView} from './BlockView';
import {breaksInWindow, windowEnd, type Anchor, type BlockBox, type LineBox, type PageStart} from './domain/paging';
import {wordCount, type Block} from './model/docx';

/**
 * Counts the pages of the whole document in the background, for the page count and the
 * Pages view. It lays the document out off-screen, a window of blocks at a time, exactly
 * as the page does (same BlockView, width, height and text size), and walks the same page
 * breaks the ▶ button would. Each window is a separate render, with a pause between, so
 * the page stays responsive. Remount it (change its key) to start again.
 *
 * `seed`: pages already known from an earlier count that an edit further on cannot have
 * moved. Counting resumes at the last of them, so an edit on page 38 re-measures pages 38
 * onward, not the whole document.
 */

type Props = {
  blocks: Block[];
  width: number;
  pageH: number;
  fonts: Set<string>;
  scale: number;
  seed?: PageStart[];
  /** Called as pages are found; `done` once the end is reached. */
  onPages: (pages: PageStart[], done: boolean) => void;
};

const FIRST: PageStart = {anchor: {block: 0, offset: 0}, char: 0};

export function PageCounter({blocks, width, pageH, fonts, scale, seed, onPages}: Props): React.JSX.Element | null {
  const counts = useMemo(() => blocks.map(wordCount), [blocks]);
  // The last seeded page is measured again: its end may have moved.
  const known = seed && seed.length > 0 ? seed.slice(0, -1) : [];
  const from = seed && seed.length > 0 ? seed[seed.length - 1] : FIRST;
  const [win, setWin] = useState<{start: Anchor; budget: number}>({start: from.anchor, budget: 2500});
  const end = useMemo(() => windowEnd(counts, win.start.block, win.budget, 250), [counts, win]);
  const window = useMemo(() => blocks.slice(win.start.block, end), [blocks, win.start.block, end]);

  const pages = useRef<PageStart[]>([...known, from]);
  const frames = useRef<Array<{top: number; height: number} | undefined>>([]);
  const lines = useRef<Array<LineBox[] | undefined>>([]);
  const settle = useRef<ReturnType<typeof setTimeout> | null>(null);
  const finished = useRef(false);

  useEffect(() => {
    frames.current = [];
    lines.current = [];
  }, [win]);

  useEffect(
    () => () => {
      if (settle.current) {
        clearTimeout(settle.current);
      }
    },
    [],
  );

  const boxes = (): Array<BlockBox | undefined> =>
    window.map((b, i) => {
      const f = frames.current[i];
      const l = lines.current[i];
      if (!f || (b.type === 'p' && b.runs.length > 0 && !l)) {
        return undefined;
      }
      return {top: f.top, height: f.height, lines: l, forced: b.type === 'p' && !!b.pb};
    });

  const measure = () => {
    if (settle.current) {
      clearTimeout(settle.current);
    }
    settle.current = setTimeout(() => {
      if (finished.current) {
        return;
      }
      const last = end >= blocks.length;
      const r = breaksInWindow(boxes(), win.start, pageH, last);
      if (!r) {
        return;
      }
      const before = pages.current.length;
      pages.current = [...pages.current, ...r.pages];
      if (!r.next) {
        finished.current = true;
        onPages(pages.current, true);
        return;
      }
      onPages(pages.current, false);
      // No page ended inside this window: it was shorter than a page. Take a bigger one.
      const grew = pages.current.length > before;
      // Pause between windows so the visible page keeps responding.
      settle.current = setTimeout(() => setWin({start: r.next!, budget: grew ? 2500 : win.budget * 2}), 80);
    }, 80);
  };

  if (width <= 0 || pageH <= 0 || blocks.length === 0) {
    return null;
  }
  return (
    <View pointerEvents="none" style={[styles.offscreen, {width}]}>
      {/* Laid out, never drawn: it is far off to the left of the screen. */}
      <View key={`${win.start.block}:${win.start.offset}:${win.budget}`} style={styles.column}>
        {window.map((b, i) => (
          <BlockView
            key={win.start.block + i}
            block={b}
            fonts={fonts}
            scale={scale}
            onFrame={(e: LayoutChangeEvent) => {
              const {y, height} = e.nativeEvent.layout;
              frames.current[i] = {top: y, height};
              measure();
            }}
            onLines={l => {
              lines.current[i] = l;
              measure();
            }}
          />
        ))}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  offscreen: {position: 'absolute', left: -20000, top: 0, opacity: 0},
  column: {position: 'absolute', left: 0, right: 0, top: 0},
});
