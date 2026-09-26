import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
  PixelRatio,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  View,
  findNodeHandle,
  type GestureResponderEvent,
  type LayoutChangeEvent,
} from 'react-native';
import {PluginManager, RattaFileSelector} from 'sn-plugin-lib';
import {BlockView} from './BlockView';
import {
  applyOps,
  comparePos,
  deletionRange,
  expectedTexts,
  formatOps,
  rangesBetween,
  styleOps,
  textEditProblem,
  wordAround,
  type FormatProp,
  type Op,
  type Pos,
  type Range,
  type StyleKind,
} from './domain/edits';
import {anchorAfter, findBreak, windowEnd, type Anchor, type BlockBox, type Break, type LineBox} from './domain/paging';
import {outline, paragraphText, wordCount, type DocxDocument, type ParagraphBlock} from './model/docx';
import {ensureFileReadPermission, ensureFileWritePermission} from './pluginPermissions';
import {Docx, DocxText, errorText, log, nativeBuild} from './services/native';

/**
 * Read a Word document a page at a time (M1) and mark it up (M2).
 *
 * Paging: the blocks from the page's anchor are drawn as one column, shifted up by the
 * anchor's offset, inside a clipped page. Once every block reports its frame (and every
 * paragraph its lines), findBreak picks where the next page starts — at a line boundary —
 * and a white mask hides the part-line below it. ▶ moves the anchor there; ◀ goes back
 * through the pages already seen.
 *
 * Editing: the pen selects words (the paragraph's own TextView reports the character under
 * the pen, via DocxText); a tap places a caret, a double tap selects a word. Buttons turn
 * the selection into edits (domain/edits), one undo step per button press. The page shows
 * the original with the first `cursor` steps applied — Undo/Redo move the cursor — and Save
 * hands the same edits to the native writer, which saves a verified copy beside the
 * original. The original file is never changed.
 */

const HEADER_H = 64;
const BAR_H = 60;
const PAD = 16;
const START: Anchor = {block: 0, offset: 0};
/** Pen travel below this (dp) is a tap: it selects the word under the pen. */
const TAP_SLOP = 12;
/** A second tap this soon, this close, selects the word under it. */
const DOUBLE_TAP_MS = 500;

type Frame = {x: number; top: number; height: number};
type Selection = {from: Pos; to: Pos};
type Typing = {mode: 'insert'; at: Pos} | {mode: 'replace'; range: Range};

export function Reader(): React.JSX.Element {
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [anchor, setAnchor] = useState<Anchor>(START);
  const [history, setHistory] = useState<Anchor[]>([]);
  const [status, setStatus] = useState('');
  const [busy, setBusy] = useState(false);
  const [contents, setContents] = useState(false);
  const [pageH, setPageH] = useState(0);
  const [measured, setMeasured] = useState<{key: string; brk: Break} | null>(null);
  const [edits, setEdits] = useState<{steps: Op[][]; cursor: number}>({steps: [], cursor: 0});
  const [saved, setSaved] = useState<{cursor: number; dest: string | null}>({cursor: 0, dest: null});
  const [selection, setSelection] = useState<Selection | null>(null);
  const [caret, setCaret] = useState<Pos | null>(null);
  const [caretBox, setCaretBox] = useState<{key: string; x: number; top: number; height: number} | null>(null);
  const [typing, setTyping] = useState<Typing | null>(null);
  const [input, setInput] = useState('');
  const [discardArmed, setDiscardArmed] = useState(false);
  const started = useRef(false);

  // Ask for write permission before the first log line: refused writes are why the probe
  // lost two logs. Saving needs it too.
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

  const applied = useMemo(() => edits.steps.slice(0, edits.cursor).flat(), [edits]);
  const blocks = useMemo(() => (doc ? applyOps(doc.blocks, applied) : []), [doc, applied]);
  const dirty = edits.cursor !== saved.cursor;
  const readOnly = !!doc && doc.report.trackedChanges > 0;

  const counts = useMemo(() => blocks.map(wordCount), [blocks]);
  const end = useMemo(() => (doc ? windowEnd(counts, anchor.block) : 0), [doc, counts, anchor.block]);
  const window = useMemo(() => blocks.slice(anchor.block, end), [blocks, anchor.block, end]);

  // Measurements belong to one page (window + offset + page height + edits). A new page
  // starts them afresh, and a break computed for another page is never used.
  const pageKey = `${doc?.path}:${anchor.block}:${anchor.offset}:${end}:${pageH}:${edits.cursor}`;
  const measuredFor = useRef('');
  const frames = useRef<Array<Frame | undefined>>([]);
  const lines = useRef<Array<LineBox[] | undefined>>([]);
  const textFrames = useRef<Array<{x: number; y: number} | undefined>>([]);
  const textRefs = useRef<Array<Text | null>>([]);
  if (measuredFor.current !== pageKey) {
    measuredFor.current = pageKey;
    frames.current = [];
    lines.current = [];
    textFrames.current = [];
  }
  const brk: Break = measured?.key === pageKey ? measured.brk : {kind: 'pending'};
  const atEnd = !doc || (brk.kind === 'after' && end >= blocks.length);

  /** A block is measured once it has a frame and, if it is a paragraph with text, its lines. */
  const boxesNow = useCallback(
    (): Array<BlockBox | undefined> =>
      window.map((b, i) => {
        const f = frames.current[i];
        const l = lines.current[i];
        if (!f || (b.type === 'p' && b.runs.length > 0 && !l)) {
          return undefined;
        }
        return {top: f.top, height: f.height, lines: l};
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
    const {x, y, height} = e.nativeEvent.layout;
    frames.current[i] = {x, top: y, height};
    remeasure();
  };

  const onLines = (i: number) => (l: LineBox[]) => {
    lines.current[i] = l;
    remeasure();
  };

  const onTextFrame = (i: number) => (e: LayoutChangeEvent) => {
    const {x, y} = e.nativeEvent.layout;
    textFrames.current[i] = {x, y};
  };

  // ---------------------------------------------------------------- open / close

  const open = async () => {
    if (dirty && !discardArmed) {
      setDiscardArmed(true);
      setStatus('You have unsaved changes. Tap Open again to discard them, or Save first.');
      return;
    }
    setDiscardArmed(false);
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
      setEdits({steps: [], cursor: 0});
      setSaved({cursor: 0, dest: null});
      setSelection(null);
      setCaret(null);
      setTyping(null);
      const r = opened.report;
      setStatus(
        r.trackedChanges > 0
          ? 'This document has tracked changes, so it is read-only here. Deleted text is hidden.'
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

  const close = () => PluginManager.closePluginView();

  // ---------------------------------------------------------------- paging

  const next = () => {
    if (!doc || atEnd) {
      return;
    }
    const to = anchorAfter(anchor, boxesNow(), brk);
    if (!to || to.block >= blocks.length) {
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

  // ---------------------------------------------------------------- pen selection

  type Hit = {para: number; offset: number; char: number};

  /** The paragraph under a viewport point (by frame), then the character (natively). */
  const hit = async (x: number, y: number): Promise<Hit | string> => {
    const colY = y + anchor.offset;
    let best: {i: number; d: number} | null = null;
    window.forEach((b, i) => {
      const f = frames.current[i];
      if (b.type !== 'p' || !f) {
        return;
      }
      const d = colY < f.top ? f.top - colY : colY > f.top + f.height ? colY - (f.top + f.height) : 0;
      if (!best || d < best.d) {
        best = {i, d};
      }
    });
    if (!best) {
      return 'no paragraph on this page';
    }
    const i = (best as {i: number}).i;
    const p = window[i] as ParagraphBlock;
    const f = frames.current[i]!;
    const tf = textFrames.current[i] ?? {x: 0, y: 0};
    const tag = findNodeHandle(textRefs.current[i] ?? null);
    if (!DocxText || tag === null) {
      return 'text lookup unavailable';
    }
    if (p.runs.length === 0) {
      return {para: p.index, offset: 0, char: 0};
    }
    const scale = PixelRatio.get();
    const lx = Math.max(0, x - f.x - tf.x) * scale;
    const ly = Math.max(0, Math.min(f.height - 1, colY - f.top - tf.y)) * scale;
    const res = await DocxText.offsetAt(tag, lx, ly);
    if (res.error !== undefined) {
      return `${res.error} (${res.viewClass})`;
    }
    return {para: p.index, offset: res.offset!, char: res.char!};
  };

  const textOf = (para: number): string => {
    const b = blocks.find(x => x.type === 'p' && x.index === para) as ParagraphBlock | undefined;
    return b ? paragraphText(b) : '';
  };

  const penFrom = useRef<{x: number; y: number} | null>(null);
  const penTo = useRef<{x: number; y: number} | null>(null);
  const selecting = useRef(false);
  const lastTap = useRef<{x: number; y: number; at: number} | null>(null);

  const release = async () => {
    const a = penFrom.current;
    const b = penTo.current;
    penFrom.current = null;
    penTo.current = null;
    if (!a || !b || selecting.current || readOnly) {
      return;
    }
    selecting.current = true;
    try {
      const tap = Math.hypot(b.x - a.x, b.y - a.y) < TAP_SLOP;
      const now = Date.now();
      const prev = lastTap.current;
      const doubleTap = tap && !!prev && now - prev.at < DOUBLE_TAP_MS && Math.hypot(a.x - prev.x, a.y - prev.y) < TAP_SLOP * 2;
      lastTap.current = tap && !doubleTap ? {x: a.x, y: a.y, at: now} : null;
      const ha = await hit(a.x, a.y);
      const hb = tap ? ha : await hit(b.x, b.y);
      if (typeof ha === 'string' || typeof hb === 'string') {
        setStatus(typeof ha === 'string' ? ha : (hb as string));
        return;
      }
      if (tap && !doubleTap) {
        // One tap: a caret in the gap nearest the pen.
        setSelection(null);
        setTyping(null);
        setCaret({para: ha.para, offset: ha.offset});
        setStatus('');
        return;
      }
      // Order by the character under each end, then snap both ends to whole words.
      const [s, e] = comparePos({para: ha.para, offset: ha.char}, {para: hb.para, offset: hb.char}) <= 0 ? [ha, hb] : [hb, ha];
      const ws = wordAround(textOf(s.para), s.char, 'right');
      const we = wordAround(textOf(e.para), e.char, 'left');
      const from = {para: s.para, offset: ws?.start ?? s.offset};
      const to = {para: e.para, offset: we?.end ?? e.offset};
      if (comparePos(from, to) >= 0) {
        setSelection(null);
        return;
      }
      setSelection({from, to});
      setCaret(null);
      setTyping(null);
      setStatus('');
    } catch (error) {
      setStatus(`Selection failed: ${errorText(error)}`);
      log(`selection failed: ${errorText(error)}`);
    } finally {
      selecting.current = false;
    }
  };

  const point = (e: GestureResponderEvent) => ({x: e.nativeEvent.locationX, y: e.nativeEvent.locationY});

  const selectedIn = (b: (typeof window)[number]): {start: number; end: number} | null => {
    if (!selection || b.type !== 'p') {
      return null;
    }
    const r = rangesBetween([b], selection.from, selection.to)[0];
    return r ? {start: r.start, end: r.end} : null;
  };

  // ---------------------------------------------------------------- edits

  /** One button press = one undo step, however many paragraphs it touches. */
  const commit = (ops: Op[], label: string) => {
    if (ops.length === 0) {
      return;
    }
    setEdits(({steps, cursor}) => ({steps: [...steps.slice(0, cursor), ops], cursor: cursor + 1}));
    // An undone save point can no longer be reached by redo.
    setSaved(sv => (sv.cursor > edits.cursor ? {...sv, cursor: -1} : sv));
    log(`${label}: ${JSON.stringify(ops)}`);
  };

  const format = (prop: FormatProp, label: string) => {
    if (!selection) {
      return;
    }
    commit(formatOps(blocks, rangesBetween(blocks, selection.from, selection.to), prop), label);
  };

  const style = (kind: StyleKind, label: string) => {
    if (!selection) {
      return;
    }
    commit(styleOps(rangesBetween(blocks, selection.from, selection.to), kind), label);
  };

  const paragraph = (para: number) => blocks.find(b => b.type === 'p' && b.index === para) as ParagraphBlock | undefined;

  const remove = () => {
    if (!selection) {
      return;
    }
    const ranges = rangesBetween(blocks, selection.from, selection.to);
    const ops: Op[] = [];
    for (const r of ranges) {
      const p = paragraph(r.para);
      const problem = p && textEditProblem(p, r.start, r.end);
      if (!p || problem) {
        setStatus(problem ?? 'Paragraph not found.');
        return;
      }
      const d = deletionRange(paragraphText(p), r.start, r.end);
      ops.push({op: 'text', para: r.para, start: d.start, end: d.end, text: ''});
    }
    commit(ops, 'delete');
    setSelection(null);
    setCaret(ranges.length ? {para: ranges[0].para, offset: ops[0].op === 'text' ? ops[0].start : 0} : null);
  };

  const startReplace = () => {
    if (!selection) {
      return;
    }
    const ranges = rangesBetween(blocks, selection.from, selection.to);
    if (ranges.length !== 1) {
      setStatus('Replace works inside one paragraph. Select less, or use Delete.');
      return;
    }
    const r = ranges[0];
    const p = paragraph(r.para);
    const problem = p ? textEditProblem(p, r.start, r.end) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    setInput(paragraphText(p!).slice(r.start, r.end));
    setTyping({mode: 'replace', range: r});
  };

  const startInsert = () => {
    if (!caret) {
      return;
    }
    const p = paragraph(caret.para);
    const problem = p ? textEditProblem(p, caret.offset, caret.offset) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    setInput('');
    setTyping({mode: 'insert', at: caret});
  };

  const finishTyping = () => {
    if (!typing) {
      return;
    }
    // One line of text: line breaks and other control characters become spaces.
    const text = input.replace(/[\r\n\u0000-\u0008\u000b-\u001f\ufffc]/g, ' ');
    const r = typing.mode === 'insert' ? {para: typing.at.para, start: typing.at.offset, end: typing.at.offset} : typing.range;
    const current = paragraphText(paragraph(r.para)!).slice(r.start, r.end);
    if (text !== current) {
      commit([{op: 'text', para: r.para, start: r.start, end: r.end, text}], typing.mode);
    }
    setTyping(null);
    setSelection(null);
    setInput('');
    setCaret({para: r.para, offset: r.start + (text !== current ? text.length : r.end - r.start)});
  };

  const cancelTyping = () => {
    setTyping(null);
    setInput('');
  };

  const undo = () => {
    setEdits(e => ({...e, cursor: Math.max(0, e.cursor - 1)}));
    setCaret(null);
    setSelection(null);
  };
  const redo = () => {
    setEdits(e => ({...e, cursor: Math.min(e.steps.length, e.cursor + 1)}));
    setCaret(null);
    setSelection(null);
  };

  const save = async () => {
    if (!doc || !Docx) {
      return;
    }
    setBusy(true);
    setStatus('Saving…');
    try {
      if (!(await ensureFileWritePermission())) {
        setStatus('Saving needs file write permission.');
        return;
      }
      const res = await Docx.save(doc.path, applied, saved.dest ?? '', expectedTexts(doc.blocks, applied));
      setSaved({cursor: edits.cursor, dest: res.dest});
      setDiscardArmed(false);
      setStatus(`Saved as ${res.name} (next to the original, which is unchanged).`);
    } catch (error) {
      setStatus(`Not saved: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  // Where to draw the caret: asked of the paragraph's own layout once the page is measured.
  useEffect(() => {
    if (!caret || brk.kind === 'pending' || !DocxText) {
      return;
    }
    const i = window.findIndex(b => b.type === 'p' && b.index === caret.para);
    const f = frames.current[i];
    const tag = i >= 0 ? findNodeHandle(textRefs.current[i] ?? null) : null;
    if (i < 0 || !f || tag === null) {
      return;
    }
    const tf = textFrames.current[i] ?? {x: 0, y: 0};
    const key = pageKey;
    DocxText.caretRect(tag, caret.offset).then(r => {
      if (r.error !== undefined || measuredFor.current !== key) {
        return;
      }
      const scale = PixelRatio.get();
      setCaretBox({key, x: f.x + tf.x + r.x! / scale, top: f.top + tf.y + r.top! / scale, height: (r.bottom! - r.top!) / scale});
    });
    // frames/refs are read at call time; brk marks "measured".
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [caret, brk, pageKey]);

  // ---------------------------------------------------------------- render

  const visible = brk.kind === 'at' ? Math.max(0, brk.top - anchor.offset) : pageH;
  const progress = doc && blocks.length > 0 ? Math.round((anchor.block / blocks.length) * 100) : 0;
  const headings = useMemo(() => outline(blocks), [blocks]);

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
          {doc ? `${dirty ? '• ' : ''}${doc.name.replace(/\.docx$/i, '')}` : 'DOCX'}
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
      <View style={styles.bar}>
        {doc && typing && !contents ? (
          <View style={styles.typing}>
            <TextInput
              style={styles.input}
              value={input}
              onChangeText={setInput}
              autoFocus
              allowFontScaling={false}
              returnKeyType="done"
              onSubmitEditing={finishTyping}
              placeholder={typing.mode === 'insert' ? 'Type text to insert…' : 'Replacement text…'}
              placeholderTextColor="#666"
            />
            {button('OK', finishTyping)}
            {button('Cancel', cancelTyping)}
          </View>
        ) : doc && selection && !contents ? (
          <ScrollView horizontal style={styles.actions} contentContainerStyle={styles.actionsInner}>
            {button('Highlight', () => format('h', 'highlight'))}
            {button('Bold', () => format('b', 'bold'))}
            {button('Italic', () => format('i', 'italic'))}
            {button('Underline', () => format('u', 'underline'))}
            {button('H1', () => style('heading1', 'heading 1'))}
            {button('H2', () => style('heading2', 'heading 2'))}
            {button('Body', () => style('normal', 'body text'))}
            {button('Delete', remove)}
            {button('Replace', startReplace)}
            {button('✕', () => setSelection(null))}
          </ScrollView>
        ) : doc && caret && !contents ? (
          <View style={styles.actions}>
            <View style={styles.actionsInner}>
              <Text allowFontScaling={false} style={styles.caretHint}>
                {'Caret placed.'}
              </Text>
              {button('Type', startInsert)}
              {button('✕', () => setCaret(null))}
            </View>
          </View>
        ) : (
          <Text allowFontScaling={false} style={styles.statusText} numberOfLines={2}>
            {status || (doc && !readOnly ? 'Drag across words to select · tap for a caret · double-tap a word.' : '')}
          </Text>
        )}
        {doc && !readOnly ? (
          <View style={styles.editButtons}>
            {button('Undo', undo, edits.cursor === 0)}
            {button('Redo', redo, edits.cursor === edits.steps.length)}
            {button('Save', save, !dirty)}
          </View>
        ) : null}
      </View>
      <View style={styles.pageArea} onLayout={e => setPageH(Math.floor(e.nativeEvent.layout.height) - PAD * 2)}>
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
            <View key={pageKey} style={[styles.column, {top: -anchor.offset}]}>
              {window.map((b, i) => (
                <BlockView
                  key={anchor.block + i}
                  block={b}
                  selection={selectedIn(b)}
                  onFrame={onFrame(i)}
                  onLines={onLines(i)}
                  onTextFrame={onTextFrame(i)}
                  textRef={t => {
                    textRefs.current[i] = t;
                  }}
                />
              ))}
              {caret && caretBox?.key === pageKey ? (
                <View pointerEvents="none" style={[styles.caret, {left: caretBox.x - 1, top: caretBox.top, height: caretBox.height}]} />
              ) : null}
            </View>
            <View style={[styles.mask, {top: visible}]} />
            {/* Owns the pen, so nothing inks and no text handles the touch itself. */}
            <View
              style={StyleSheet.absoluteFill}
              onStartShouldSetResponder={() => !readOnly}
              onMoveShouldSetResponder={() => !readOnly}
              onResponderTerminationRequest={() => false}
              onResponderGrant={e => {
                penFrom.current = point(e);
                penTo.current = penFrom.current;
              }}
              onResponderMove={e => {
                penTo.current = point(e);
              }}
              onResponderRelease={release}
              onResponderTerminate={release}
            />
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
  bar: {
    height: BAR_H,
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: PAD,
    borderBottomWidth: 1,
    borderColor: '#000',
  },
  actions: {flex: 1},
  actionsInner: {alignItems: 'center'},
  editButtons: {flexDirection: 'row', marginLeft: 8},
  typing: {flex: 1, flexDirection: 'row', alignItems: 'center'},
  input: {flex: 1, height: 44, borderWidth: 1, borderColor: '#000', paddingHorizontal: 10, fontSize: 18, color: '#000'},
  caretHint: {color: '#000', fontSize: 15},
  caret: {position: 'absolute', width: 3, backgroundColor: '#000'},
  statusText: {flex: 1, color: '#000', fontSize: 15},
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
