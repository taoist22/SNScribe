import React from 'react';
import {Image, StyleSheet, Text, View, type LayoutChangeEvent, type TextLayoutEventData, type NativeSyntheticEvent} from 'react-native';
import {markSelection, splitRuns} from './domain/edits';
import {shownFont} from './domain/fonts';
import {OBJECT, type Block, type ParagraphBlock, type Run} from './model/docx';
import type {LineBox} from './domain/paging';

/**
 * One block of the document. A paragraph is ONE Text with a nested Text per run: spike S1
 * measured ~0.19 s a screen this way against ~4.4 s for a Text per word, and the pen will
 * find characters through the paragraph's own layout (spike S1b).
 *
 * Inline objects keep their single character (so offsets match the model) but show as a
 * symbol: ▣ image, * note reference, ◇ other object.
 *
 * Grayscale only: highlight is light grey, links are underlined. Tracked changes as Word
 * shows them, in grey: insertions underlined, deletions struck through (shown before the
 * character they sit at, though they are not part of the text — see shownExtras).
 */

type Props = {
  block: Block;
  /** Selected characters of this paragraph, drawn inverted. */
  selection?: {start: number; end: number} | null;
  /** Misspelled words: light grey and underlined (the screen can't draw Word's red squiggle). */
  spell?: Array<{start: number; end: number}>;
  onFrame: (e: LayoutChangeEvent) => void;
  onLines?: (lines: LineBox[]) => void;
  /** The paragraph's Text: its native view is asked which character is under the pen. */
  textRef?: (t: Text | null) => void;
  /** The paragraph Text's position inside the block frame (list rows shift it right). */
  onTextFrame?: (e: LayoutChangeEvent) => void;
  /** Font families loaded on this device; others show in their loaded stand-in (domain/fonts), else the default font. */
  fonts?: Set<string>;
  /** Text size factor chosen by the reader (A− / A+); 1 = the document's sizes. */
  scale?: number;
  /** The text column's width: pictures wider than it are shown smaller. */
  width?: number;
};

const SYMBOL: Record<string, string> = {image: '▣', note: '*', object: '◇', ink: '✎'};
/** The list number's box: the number sits in the list's hanging space, as in Word. */
const LIST_LABEL_W = 28;

function runText(r: Run): string {
  return r.obj ? SYMBOL[r.obj] ?? '◇' : r.t.split(OBJECT).join('◇');
}

function sizeFor(p: ParagraphBlock): number {
  switch (p.kind) {
    case 'title':
      return 34;
    case 'subtitle':
      return 24;
    case 'heading':
      return p.level <= 1 ? 30 : p.level === 2 ? 26 : p.level === 3 ? 23 : 21;
    default:
      return 20;
  }
}

/** Word sizes (half-points) on screen: 11 pt body text ≈ 20 dp, as before sizes were read. */
const dpFor = (halfPoints: number) => Math.max(10, Math.min(64, Math.round(halfPoints * 0.91)));
/** Word lengths in twips (1/20 pt) on screen, on the same scale as sizes. */
const dpForTwips = (twips: number, scale: number) => Math.max(0, Math.min(120, Math.round((twips / 20) * 1.82 * scale)));

/**
 * Characters the paragraph's TextView holds before its text: a first-line indent is drawn
 * as an inline spacer, which Android counts as one character. Offsets from the native
 * layout (hit tests, carets, line moves) are shifted by this.
 */
export function leadChars(b: Block): number {
  return b.type === 'p' && (b.first ?? 0) > 0 && b.runs.length > 0 ? 1 : 0;
}

type Piece = Run & {sel?: boolean; del?: boolean; sp?: boolean; caret?: boolean};


/** The caret drawn in the text while typing, so it shows in the same screen update as the typed letters. */
const CARET = '|';

/** The pieces with the caret before model character `caret` (after deleted text shown there). */
function withCaret(ps: Piece[], caret: number | undefined): Piece[] {
  if (caret === undefined) {
    return ps;
  }
  const out: Piece[] = [];
  let off = 0;
  let placed = false;
  for (const r of ps) {
    if (!placed && !r.del && caret < off + r.t.length) {
      const k = caret - off;
      if (k > 0) {
        out.push({...r, t: r.t.slice(0, k)});
      }
      out.push({t: CARET, caret: true, sz: r.sz});
      out.push({...r, t: r.t.slice(k)});
      placed = true;
    } else {
      out.push(r);
    }
    if (!r.del) {
      off += r.t.length;
    }
  }
  if (!placed) {
    const last = [...ps].reverse().find(r => !r.del);
    out.push({t: CARET, caret: true, sz: last?.sz});
  }
  return out;
}

/** The paragraph's runs with selection and misspellings marked, and deleted text placed where it sits. */
function pieces(p: ParagraphBlock, selection: {start: number; end: number} | null, spell?: Array<{start: number; end: number}>): Piece[] {
  let marked: Piece[] = markSelection(p.runs, selection);
  if (spell?.length) {
    let off = 0;
    marked = (splitRuns(marked, spell.flatMap(r => [r.start, r.end])) as Piece[]).map(r => {
      const s0 = off;
      off += r.t.length;
      return spell.some(x => s0 >= x.start && off <= x.end) && r.t.length > 0 ? {...r, sp: true} : r;
    });
  }
  const dels = (p.revs ?? []).filter(v => v.kind === 'del' && v.runs?.length).sort((a, b) => (a.at ?? 0) - (b.at ?? 0));
  if (dels.length === 0) {
    return marked;
  }
  const out: Piece[] = [];
  let offset = 0;
  let d = 0;
  for (const r of splitRuns(marked, dels.map(v => v.at ?? 0)) as Piece[]) {
    while (d < dels.length && (dels[d].at ?? 0) <= offset) {
      out.push(...dels[d++].runs!.map(x => ({...x, del: true})));
    }
    out.push(r);
    offset += r.t.length;
  }
  while (d < dels.length) {
    out.push(...dels[d++].runs!.map(x => ({...x, del: true})));
  }
  return out;
}

/** A picture's size on screen: Word's size (EMU) on the same scale as text, no wider than `maxW`. */
function pictureSize(r: Run, scale: number, maxW: number): {width: number; height: number} {
  const toDp = (emu: number) => (emu / 12700) * 1.82 * scale;
  let w = r.cx ? toDp(r.cx) : 200;
  let h = r.cy ? toDp(r.cy) : 150;
  if (w > maxW) {
    h = (h * maxW) / w;
    w = maxW;
  }
  return {width: Math.max(8, Math.round(w)), height: Math.max(8, Math.round(h))};
}

function ParagraphView({p, selection, spell, onFrame, onLines, textRef, onTextFrame, fonts, scale = 1, width = 700}: Omit<Props, 'block'> & {p: ParagraphBlock}) {
  // The document's own sizes when it has them, times the reader's text size; the line
  // height follows the largest.
  // Text without a size of its own takes its style's (bs), else the kind's default.
  const base = p.bs ? Math.round(dpFor(p.bs) * scale) : Math.round(sizeFor(p) * scale);
  const runSize = (r: Run) => (r.sz ? Math.round(dpFor(r.sz) * scale) : base);
  const family = (r: Run) => r.f ?? p.bf;
  // What the paragraph's style gives text that does not say otherwise.
  const styleBold = (p.kind !== 'body' && p.kind !== 'subtitle') || !!p.sb;
  const styleItalic = p.kind === 'subtitle' || !!p.quote || !!p.si;
  const size = p.runs.length ? Math.max(...p.runs.map(runSize)) : base;
  // The document's line spacing when it has one: 'auto' in 240ths of a line, else twips.
  const lineHeight =
    p.line === undefined
      ? Math.round(size * 1.45)
      : p.lineRule === 'auto' || p.lineRule === undefined
      ? Math.max(size, Math.round(size * 1.2 * (p.line / 240)))
      : p.lineRule === 'atLeast'
      ? Math.max(Math.round(size * 1.2), dpForTwips(p.line, scale))
      : Math.max(8, dpForTwips(p.line, scale));
  const heading = p.kind !== 'body';
  // Word allows negative indents (text pulled into the page margin); the screen has no
  // margin to pull into, so they start at the left edge instead of off it.
  // A block quote with no indent of its own shows indented anyway, as Word's Quote style does.
  const indent = Math.max(0, Math.min(160, Math.round(((p.quote && !p.indent ? 720 : p.indent) / 20) * scale)));
  // A picture makes its line as tall as itself: a fixed line height would draw it over the lines around it.
  const hasPicture = p.runs.some(r => r.obj === 'image' && r.src);
  const maxPicture = Math.max(60, width - indent - (p.list !== undefined ? LIST_LABEL_W : 0) - 4);
  const text = (
    <Text
      ref={textRef}
      onLayout={onTextFrame}
      allowFontScaling={false}
      style={[
        styles.text,
        hasPicture ? {fontSize: size, textAlign: p.align} : {fontSize: size, lineHeight, textAlign: p.align},
        (heading && p.kind !== 'subtitle') || p.sb ? styles.bold : null,
        p.kind === 'subtitle' || p.quote || p.si ? styles.italic : null,
        p.list !== undefined ? styles.flex : null,
        p.runs.length === 0 ? {minHeight: lineHeight} : null,
      ]}
      onTextLayout={(e: NativeSyntheticEvent<TextLayoutEventData>) =>
        onLines?.(e.nativeEvent.lines.map(l => ({y: l.y, height: l.height, len: l.text.length})))
      }>
      {leadChars(p) ? <View key="first-line" style={{width: dpForTwips(p.first!, scale), height: 1}} /> : null}
      {withCaret(pieces(p, selection ?? null, spell), p.caret).map((r, i) =>
        r.obj === 'note' && r.nn ? (
          // The footnote's number, raised: one character in the TextView, like the mark it replaces.
          <View key={i} style={[styles.noteBox, {height: Math.round(runSize(r) * 0.95)}, r.sel ? styles.noteSelected : null]}>
            <Text allowFontScaling={false} style={[styles.noteNumber, {fontSize: Math.round(runSize(r) * 0.62), lineHeight: Math.round(runSize(r) * 0.7)}, r.sel ? styles.noteNumberSelected : null]}>
              {r.nn}
            </Text>
          </View>
        ) : r.obj === 'image' && r.src ? (
          // One character in the TextView, like the placeholder it replaces: offsets stay the same.
          <Image key={i} source={{uri: `file://${r.src}`}} style={[pictureSize(r, scale, maxPicture), r.sel ? styles.pictureSelected : null]} resizeMode="contain" />
        ) : r.caret ? (
          // Pulled together (negative spacing) so the text beside it barely moves.
          <Text key={i} style={[styles.caret, {fontSize: runSize(r), letterSpacing: -Math.round(runSize(r) * 0.18)}]}>
            {CARET}
          </Text>
        ) : (
        <Text
          key={i}
          style={[
            // Bold and italic are set on every run, never inherited from the paragraph: a run
            // that names its own font starts from a regular face on Android, so a heading's bold
            // (from its style, on the paragraph) did not show on Times New Roman text (CT).
            (r.b ?? styleBold) ? styles.bold : styles.notBold,
            (r.i ?? styleItalic) ? styles.italic : styles.notItalic,
            r.h ? styles.highlight : null,
            r.del ? styles.deleted : r.rv ? styles.inserted : null,
            r.sp ? styles.misspelled : null,
            r.u || r.s || r.del || r.rv || r.sp
              ? {textDecorationLine: (r.u || r.rv || r.sp) && (r.s || r.del) ? 'underline line-through' : r.u || r.rv || r.sp ? 'underline' : 'line-through'}
              : null,
            {fontSize: r.sup || r.sub ? Math.round(runSize(r) * 0.65) : runSize(r)},
            shownFont(family(r), fonts) ? {fontFamily: shownFont(family(r), fonts)} : null,
            r.sel ? styles.selected : null,
          ]}>
          {runText(r)}
        </Text>
        ),
      )}
    </Text>
  );
  // The document's own space before and after when it has them.
  const spacing = {
    marginTop: p.before !== undefined ? dpForTwips(p.before, scale) : heading ? Math.round(size * 0.6) : 0,
    marginBottom: p.after !== undefined ? dpForTwips(p.after, scale) : heading ? Math.round(size * 0.3) : Math.round(12 * scale),
  };
  if (p.list === undefined) {
    return (
      <View style={[spacing, {marginLeft: indent}]} onLayout={onFrame}>
        {text}
      </View>
    );
  }
  return (
    <View style={[spacing, styles.row, {marginLeft: Math.max(0, indent - LIST_LABEL_W)}]} onLayout={onFrame}>
      <Text allowFontScaling={false} style={[styles.text, styles.label, {fontSize: size, lineHeight}]}>
        {p.list}
      </Text>
      {text}
    </View>
  );
}

const sameRange = (a?: {start: number; end: number} | null, b?: {start: number; end: number} | null) =>
  a === b || (!!a && !!b && a.start === b.start && a.end === b.end) || (!a && !b);

const sameRanges = (a?: Array<{start: number; end: number}>, b?: Array<{start: number; end: number}>) =>
  a === b || ((a?.length ?? 0) === (b?.length ?? 0) && (a ?? []).every((r, i) => sameRange(r, b![i])));

/**
 * Drawn again only when what it shows changes: typing redraws the one paragraph typed in,
 * not the page. The callbacks are left out of the comparison: callers keep them meaning
 * the same thing for the same slot (Reader reads the latest state through refs).
 */
export const BlockView = React.memo(
  BlockViewInner,
  (a, b) =>
    a.block === b.block &&
    sameRange(a.selection, b.selection) &&
    sameRanges(a.spell, b.spell) &&
    a.fonts === b.fonts &&
    a.scale === b.scale &&
    a.width === b.width,
);

function BlockViewInner({block, ...rest}: Props): React.JSX.Element {
  if (block.type === 'p') {
    if (block.pb) {
      // Outside the paragraph's frame (so its frame and lines stay its own): the marker ends
      // the previous page, and the paragraph starts the next one (paging forces the break).
      return (
        <>
          <View style={styles.pageBreak}>
            <Text allowFontScaling={false} style={styles.pageBreakText}>
              {'— page break —'}
            </Text>
          </View>
          <ParagraphView p={block} {...rest} />
        </>
      );
    }
    return <ParagraphView p={block} {...rest} />;
  }
  const {onFrame} = rest;
  const title =
    block.type === 'table'
      ? `▦ Table · ${block.rows} × ${block.cols} — shown as a summary`
      : `▦ ${block.what} — not shown yet`;
  return (
    <View style={styles.locked} onLayout={onFrame}>
      <Text allowFontScaling={false} style={styles.lockedTitle}>
        {title}
      </Text>
      {block.preview ? (
        <Text allowFontScaling={false} style={styles.lockedPreview} numberOfLines={3}>
          {block.preview}
        </Text>
      ) : null}
    </View>
  );
}

const styles = StyleSheet.create({
  text: {color: '#000'},
  flex: {flex: 1},
  row: {flexDirection: 'row'},
  label: {width: LIST_LABEL_W},
  bold: {fontWeight: '700'},
  italic: {fontStyle: 'italic'},
  notBold: {fontWeight: '400'},
  notItalic: {fontStyle: 'normal'},
  highlight: {backgroundColor: '#cfcfcf'},
  selected: {backgroundColor: '#000', color: '#fff'},
  caret: {color: '#000', fontWeight: '400', fontStyle: 'normal'},
  pictureSelected: {opacity: 0.4},
  noteBox: {justifyContent: 'flex-start', paddingHorizontal: 1},
  noteSelected: {backgroundColor: '#000'},
  noteNumber: {color: '#000', fontWeight: '700'},
  noteNumberSelected: {color: '#fff'},
  inserted: {color: '#333'},
  misspelled: {backgroundColor: '#e6e6e6'},
  deleted: {color: '#777'},
  pageBreak: {height: 28, justifyContent: 'center', alignItems: 'center', borderTopWidth: 1, borderColor: '#000', borderStyle: 'dashed'},
  pageBreakText: {color: '#000', fontSize: 13},
  locked: {borderWidth: 1, borderColor: '#000', borderStyle: 'dashed', padding: 10, marginBottom: 12},
  lockedTitle: {color: '#000', fontSize: 17, fontWeight: '700'},
  lockedPreview: {color: '#333', fontSize: 16, fontStyle: 'italic', marginTop: 4},
});
