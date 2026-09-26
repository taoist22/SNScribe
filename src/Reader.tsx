import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {Pressable, ScrollView, StyleSheet, Text, View, type LayoutChangeEvent} from 'react-native';
import {PluginManager, RattaFileSelector} from 'sn-plugin-lib';
import {BlockView} from './BlockView';
import {anchorAfter, findBreak, windowEnd, type Anchor, type BlockBox, type Break, type LineBox} from './domain/paging';
import {outline, wordCount, type DocxDocument} from './model/docx';
import {ensureFileReadPermission, ensureFileWritePermission} from './pluginPermissions';
import {Docx, errorText, log, nativeBuild} from './services/native';

/**
 * Milestone 1: open a Word document and read it, a page at a time.
 *
 * Paging: the blocks from the page's anchor are drawn as one column, shifted up by the
 * anchor's offset, inside a clipped page. Once every block reports its frame (and every
 * paragraph its lines), findBreak picks where the next page starts — at a line boundary —
 * and a white mask hides the part-line below it. ▶ moves the anchor there; ◀ goes back
 * through the pages already seen.
 */

const HEADER_H = 64;
const STATUS_H = 34;
const PAD = 16;
const START: Anchor = {block: 0, offset: 0};

export function Reader(): React.JSX.Element {
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [anchor, setAnchor] = useState<Anchor>(START);
  const [history, setHistory] = useState<Anchor[]>([]);
  const [status, setStatus] = useState('');
  const [busy, setBusy] = useState(false);
  const [contents, setContents] = useState(false);
  const [pageH, setPageH] = useState(0);
  const [measured, setMeasured] = useState<{key: string; brk: Break} | null>(null);
  const started = useRef(false);

  // Ask for write permission before the first log line: refused writes are why the probe
  // lost two logs.
  useEffect(() => {
    if (started.current) {
      return;
    }
    started.current = true;
    (async () => {
      const canWrite = await ensureFileWritePermission();
      const name = await Docx?.logName().catch(() => '?');
      const refused = await log(`DOCX opened: NATIVE_BUILD=${nativeBuild()} write=${canWrite} log=${name}`);
      if (refused) {
        setStatus(`Log not written: ${refused}`);
      }
    })();
  }, []);

  const counts = useMemo(() => doc?.blocks.map(wordCount) ?? [], [doc]);
  const end = useMemo(() => (doc ? windowEnd(counts, anchor.block) : 0), [doc, counts, anchor.block]);
  const window = useMemo(() => doc?.blocks.slice(anchor.block, end) ?? [], [doc, anchor.block, end]);

  // Measurements belong to one page (window + offset + page height). A new page starts
  // them afresh, and a break computed for another page is never used.
  const pageKey = `${doc?.path}:${anchor.block}:${anchor.offset}:${end}:${pageH}`;
  const measuredFor = useRef('');
  const frames = useRef<Array<{top: number; height: number} | undefined>>([]);
  const lines = useRef<Array<LineBox[] | undefined>>([]);
  if (measuredFor.current !== pageKey) {
    measuredFor.current = pageKey;
    frames.current = [];
    lines.current = [];
  }
  const brk: Break = measured?.key === pageKey ? measured.brk : {kind: 'pending'};
  const atEnd = !doc || (brk.kind === 'after' && end >= doc.blocks.length);

  /** A block is measured once it has a frame and, if it is a paragraph with text, its lines. */
  const boxesNow = useCallback(
    (): Array<BlockBox | undefined> =>
      window.map((b, i) => {
        const f = frames.current[i];
        const l = lines.current[i];
        if (!f || (b.type === 'p' && b.runs.length > 0 && !l)) {
          return undefined;
        }
        return {...f, lines: l};
      }),
    [window],
  );

  const settle = useRef<ReturnType<typeof setTimeout> | null>(null);
  const remeasure = () => {
    if (settle.current) {
      clearTimeout(settle.current);
    }
    const key = pageKey;
    settle.current = setTimeout(() => {
      if (pageH > 0 && measuredFor.current === key) {
        setMeasured({key, brk: findBreak(boxesNow(), anchor.offset, pageH)});
      }
    }, 60);
  };

  const onFrame = (i: number) => (e: LayoutChangeEvent) => {
    const {y, height} = e.nativeEvent.layout;
    frames.current[i] = {top: y, height};
    remeasure();
  };

  const onLines = (i: number) => (l: LineBox[]) => {
    lines.current[i] = l;
    remeasure();
  };

  const open = async () => {
    setBusy(true);
    try {
      if (!(await ensureFileReadPermission())) {
        setStatus('File access was not allowed.');
        return;
      }
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0, // 1 opens the file and never returns (reference_sdk_calls_that_navigate)
        maxNum: 1,
        title: 'Open a Word document',
        rightButtonText: 'Open',
        suffixList: ['docx'],
      })) as string[] | null | undefined;
      const path = picked?.find(p => typeof p === 'string' && p.length > 0);
      if (!path) {
        return;
      }
      if (!/\.docx$/i.test(path)) {
        setStatus('Only .docx files can be opened (not .doc).');
        return;
      }
      setStatus('Opening…');
      const opened = await Docx!.open(path);
      setDoc(opened);
      setAnchor(START);
      setHistory([]);
      setContents(false);
      const r = opened.report;
      setStatus(
        r.trackedChanges > 0
          ? 'This document has tracked changes: deleted text is hidden, and it will open read-only for editing.'
          : '',
      );
      log(`opened ${opened.name}: ${opened.blocks.length} blocks in ${opened.ms} ms; ${JSON.stringify(r)}`);
    } catch (error) {
      setStatus(`Could not open: ${errorText(error)}`);
      log(`open failed: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  const next = () => {
    if (!doc || atEnd) {
      return;
    }
    const to = anchorAfter(anchor, boxesNow(), brk);
    if (!to || to.block >= doc.blocks.length) {
      return;
    }
    setHistory(h => [...h, anchor]);
    setAnchor(to);
  };

  const previous = () => {
    if (history.length === 0) {
      return;
    }
    setAnchor(history[history.length - 1]);
    setHistory(history.slice(0, -1));
  };

  const jump = (block: number) => {
    setHistory(h => [...h, anchor]);
    setAnchor({block, offset: 0});
    setContents(false);
  };

  const close = () => PluginManager.closePluginView();

  const visible = brk.kind === 'at' ? Math.max(0, brk.top - anchor.offset) : pageH;
  const progress = doc && doc.blocks.length > 0 ? Math.round((anchor.block / doc.blocks.length) * 100) : 0;
  const headings = useMemo(() => (doc ? outline(doc.blocks) : []), [doc]);

  const button = (label: string, action: () => void, disabled = false) => (
    <Pressable key={label} disabled={disabled || busy} onPress={action} style={[styles.button, disabled || busy ? styles.disabled : null]}>
      <Text allowFontScaling={false} style={styles.buttonText}>
        {label}
      </Text>
    </Pressable>
  );

  return (
    <View style={styles.root}>
      <View style={styles.header}>
        {button('Open', open)}
        <Text allowFontScaling={false} style={styles.title} numberOfLines={1}>
          {doc ? doc.name.replace(/\.docx$/i, '') : 'DOCX'}
        </Text>
        {doc ? button('Contents', () => setContents(c => !c), headings.length === 0) : null}
        {doc ? button('◀', previous, history.length === 0) : null}
        {doc ? (
          <Text allowFontScaling={false} style={styles.page}>
            {`${history.length + 1} · ${progress}%`}
          </Text>
        ) : null}
        {doc ? button('▶', next, atEnd || brk.kind === 'pending') : null}
        {button('Close', close)}
      </View>
      <View style={styles.status}>
        <Text allowFontScaling={false} style={styles.statusText} numberOfLines={1}>
          {status}
        </Text>
      </View>
      <View
        style={styles.pageArea}
        onLayout={e => setPageH(Math.floor(e.nativeEvent.layout.height) - PAD * 2)}>
        {!doc ? (
          <View style={styles.empty}>
            <Text allowFontScaling={false} style={styles.emptyText}>
              {'Tap Open to choose a Word document (.docx).'}
            </Text>
          </View>
        ) : contents ? (
          <ScrollView style={styles.contents}>
            {headings.map(h => (
              <Pressable key={h.block} onPress={() => jump(h.block)} style={styles.contentsRow}>
                <Text allowFontScaling={false} style={[styles.contentsText, {marginLeft: Math.max(0, h.level - 1) * 24}]} numberOfLines={2}>
                  {h.text}
                </Text>
              </Pressable>
            ))}
          </ScrollView>
        ) : pageH > 0 ? (
          <View style={[styles.viewport, {height: pageH}]}>
            <View key={`${anchor.block}:${anchor.offset}`} style={[styles.column, {top: -anchor.offset}]}>
              {window.map((b, i) => (
                <BlockView key={anchor.block + i} block={b} onFrame={onFrame(i)} onLines={onLines(i)} />
              ))}
            </View>
            <View style={[styles.mask, {top: visible}]} />
          </View>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {flex: 1, backgroundColor: '#fff'},
  header: {
    height: HEADER_H,
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: PAD,
    borderBottomWidth: 1,
    borderColor: '#000',
  },
  title: {flex: 1, color: '#000', fontSize: 20, fontWeight: '700', marginHorizontal: 12},
  page: {color: '#000', fontSize: 16, minWidth: 72, textAlign: 'center'},
  button: {borderWidth: 1, borderColor: '#000', borderRadius: 5, paddingVertical: 8, paddingHorizontal: 14, marginLeft: 8},
  buttonText: {color: '#000', fontSize: 16, fontWeight: '700'},
  disabled: {opacity: 0.3},
  status: {height: STATUS_H, justifyContent: 'center', paddingHorizontal: PAD},
  statusText: {color: '#000', fontSize: 14},
  pageArea: {flex: 1, padding: PAD},
  viewport: {overflow: 'hidden'},
  column: {position: 'absolute', left: 0, right: 0},
  mask: {position: 'absolute', left: 0, right: 0, bottom: 0, backgroundColor: '#fff'},
  empty: {flex: 1, alignItems: 'center', justifyContent: 'center'},
  emptyText: {color: '#000', fontSize: 20},
  contents: {flex: 1},
  contentsRow: {paddingVertical: 12, borderBottomWidth: 1, borderColor: '#ccc'},
  contentsText: {color: '#000', fontSize: 19},
});
