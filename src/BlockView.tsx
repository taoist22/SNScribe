import React from 'react';
import {StyleSheet, Text, View, type LayoutChangeEvent, type TextLayoutEventData, type NativeSyntheticEvent} from 'react-native';
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
 * Grayscale only: highlight is light grey, links are underlined.
 */

type Props = {
  block: Block;
  onFrame: (e: LayoutChangeEvent) => void;
  onLines?: (lines: LineBox[]) => void;
};

const SYMBOL: Record<string, string> = {image: '▣', note: '*', object: '◇'};
const LIST_LABEL_W = 40;

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

function ParagraphView({p, onFrame, onLines}: {p: ParagraphBlock; onFrame: Props['onFrame']; onLines?: Props['onLines']}) {
  const size = sizeFor(p);
  const lineHeight = Math.round(size * 1.45);
  const heading = p.kind !== 'body';
  // Word allows negative indents (text pulled into the page margin); the screen has no
  // margin to pull into, so they start at the left edge instead of off it.
  const indent = Math.max(0, Math.min(160, Math.round(p.indent / 20)));
  const text = (
    <Text
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
        onLines?.(e.nativeEvent.lines.map(l => ({y: l.y, height: l.height})))
      }>
      {p.runs.map((r, i) => (
        <Text
          key={i}
          style={[
            r.b ? styles.bold : null,
            r.i ? styles.italic : null,
            r.h ? styles.highlight : null,
            r.u || r.s ? {textDecorationLine: r.u && r.s ? 'underline line-through' : r.u ? 'underline' : 'line-through'} : null,
            r.sup ? {fontSize: Math.round(size * 0.65)} : null,
          ]}>
          {runText(r)}
        </Text>
      ))}
    </Text>
  );
  const spacing = {
    marginTop: heading ? Math.round(size * 0.6) : 0,
    marginBottom: heading ? Math.round(size * 0.3) : 12,
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

export function BlockView({block, onFrame, onLines}: Props): React.JSX.Element {
  if (block.type === 'p') {
    return <ParagraphView p={block} onFrame={onFrame} onLines={onLines} />;
  }
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
  locked: {borderWidth: 1, borderColor: '#000', borderStyle: 'dashed', padding: 10, marginBottom: 12},
  lockedTitle: {color: '#000', fontSize: 17, fontWeight: '700'},
  lockedPreview: {color: '#333', fontSize: 16, fontStyle: 'italic', marginTop: 4},
});
