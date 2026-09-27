import React from 'react';
import {StyleSheet, Text, View, type LayoutChangeEvent, type TextLayoutEventData, type NativeSyntheticEvent} from 'react-native';
import {markSelection, splitRuns} from './domain/edits';
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
  onFrame: (e: LayoutChangeEvent) => void;
  onLines?: (lines: LineBox[]) => void;
  /** The paragraph's Text: its native view is asked which character is under the pen. */
  textRef?: (t: Text | null) => void;
  /** The paragraph Text's position inside the block frame (list rows shift it right). */
  onTextFrame?: (e: LayoutChangeEvent) => void;
  /** Font families loaded on this device; other fonts show in the default font. */
  fonts?: Set<string>;
  /** Text size factor chosen by the reader (A− / A+); 1 = the document's sizes. */
  scale?: number;
};

const SYMBOL: Record<string, string> = {image: '▣', note: '*', object: '◇'};
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

type Piece = Run & {sel?: boolean; del?: boolean};

/** The paragraph's runs with selection marked, and deleted text placed where it sits. */
function pieces(p: ParagraphBlock, selection: {start: number; end: number} | null): Piece[] {
  const marked: Piece[] = markSelection(p.runs, selection);
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

function ParagraphView({p, selection, onFrame, onLines, textRef, onTextFrame, fonts, scale = 1}: Omit<Props, 'block'> & {p: ParagraphBlock}) {
  // The document's own sizes when it has them, times the reader's text size; the line
  // height follows the largest.
  const base = Math.round(sizeFor(p) * scale);
  const runSize = (r: Run) => (r.sz ? Math.round(dpFor(r.sz) * scale) : base);
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
  const indent = Math.max(0, Math.min(160, Math.round((p.indent / 20) * scale)));
  const text = (
    <Text
      ref={textRef}
      onLayout={onTextFrame}
      allowFontScaling={false}
      style={[
        styles.text,
        {fontSize: size, lineHeight, textAlign: p.align},
        heading && p.kind !== 'subtitle' ? styles.bold : null,
        p.kind === 'subtitle' ? styles.italic : null,
        p.list !== undefined ? styles.flex : null,
        p.runs.length === 0 ? {minHeight: lineHeight} : null,
      ]}
      onTextLayout={(e: NativeSyntheticEvent<TextLayoutEventData>) =>
        onLines?.(e.nativeEvent.lines.map(l => ({y: l.y, height: l.height, len: l.text.length})))
      }>
      {leadChars(p) ? <View key="first-line" style={{width: dpForTwips(p.first!, scale), height: 1}} /> : null}
      {pieces(p, selection ?? null).map((r, i) => (
        <Text
          key={i}
          style={[
            r.b ? styles.bold : null,
            r.i ? styles.italic : null,
            r.h ? styles.highlight : null,
            r.del ? styles.deleted : r.rv ? styles.inserted : null,
            r.u || r.s || r.del || r.rv
              ? {textDecorationLine: (r.u || r.rv) && (r.s || r.del) ? 'underline line-through' : r.u || r.rv ? 'underline' : 'line-through'}
              : null,
            {fontSize: r.sup ? Math.round(runSize(r) * 0.65) : runSize(r)},
            r.f && fonts?.has(r.f) ? {fontFamily: r.f} : null,
            r.sel ? styles.selected : null,
          ]}>
          {runText(r)}
        </Text>
      ))}
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

export function BlockView({block, ...rest}: Props): React.JSX.Element {
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
  highlight: {backgroundColor: '#cfcfcf'},
  selected: {backgroundColor: '#000', color: '#fff'},
  inserted: {color: '#333'},
  deleted: {color: '#777'},
  pageBreak: {height: 28, justifyContent: 'center', alignItems: 'center', borderTopWidth: 1, borderColor: '#000', borderStyle: 'dashed'},
  pageBreakText: {color: '#000', fontSize: 13},
  locked: {borderWidth: 1, borderColor: '#000', borderStyle: 'dashed', padding: 10, marginBottom: 12},
  lockedTitle: {color: '#000', fontSize: 17, fontWeight: '700'},
  lockedPreview: {color: '#333', fontSize: 16, fontStyle: 'italic', marginTop: 4},
});
