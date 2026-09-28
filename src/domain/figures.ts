// Pictures as figures: the edits that put a picture after a paragraph, with the figure
// label and title each paper format asks for, numbered in order —
//   APA 7:   "Figure 1" (bold) and the title (italic, title case) ABOVE the picture;
//   MLA 9:   "Fig. 1. Title" below it;
//   Chicago: "Figure 1. Title" below it.
// Figures after the new one are renumbered. (Mentions in the text, "see Figure 2", are not.)

import {paragraphText, type Block, type ParagraphBlock} from '../model/docx';
import type {CiteStyle} from './citations';
import type {Op} from './edits';

/** Picture sizes offered: a share of the text width. */
export const PICTURE_SIZES: Array<{name: string; frac: number}> = [
  {name: 'Small', frac: 0.25},
  {name: 'Medium', frac: 0.5},
  {name: 'Large', frac: 0.75},
  {name: 'Full width', frac: 1},
];

/** A picture `frac` of the text width (`maxW` EMU) wide, keeping its proportions (w:h), and no taller than `maxH` EMU. */
export function pictureAtWidth(w: number, h: number, frac: number, maxW: number, maxH: number): {cx: number; cy: number} {
  let cx = maxW * frac;
  let cy = (cx * Math.max(1, h)) / Math.max(1, w);
  if (cy > maxH) {
    cx = (cx * maxH) / cy;
    cy = maxH;
  }
  return {cx: Math.max(1, Math.round(cx)), cy: Math.max(1, Math.round(cy))};
}

/** A picture's size in EMU: its pixels at 96 dpi, shrunk (never enlarged) to fit `maxW` × `maxH` EMU. */
export function pictureEmu(pxW: number, pxH: number, maxW: number, maxH: number): {cx: number; cy: number} {
  const EMU_PER_PX = 9525;
  let cx = Math.max(1, pxW) * EMU_PER_PX;
  let cy = Math.max(1, pxH) * EMU_PER_PX;
  const k = Math.min(1, maxW / cx, maxH / cy);
  cx = Math.round(cx * k);
  cy = Math.round(cy * k);
  return {cx: Math.max(1, cx), cy: Math.max(1, cy)};
}

/** A figure label paragraph in `style`, and its number: "Figure 3" (APA), "Fig. 3." (MLA), "Figure 3." (Chicago). */
function labelNumber(text: string, style: CiteStyle, kind: 'figure' | 'table' = 'figure'): {n: number; at: number; len: number} | null {
  const re =
    kind === 'table'
      ? style === 'apa' || style === 'mla'
        ? /^(Table )(\d+)\s*$/
        : /^(Table )(\d+)\./
      : style === 'apa'
      ? /^(Figure )(\d+)\s*$/
      : style === 'mla'
      ? /^(Fig\. )(\d+)\./
      : /^(Figure )(\d+)\./;
  const m = re.exec(text.trim());
  if (!m) {
    return null;
  }
  const lead = text.length - text.trimStart().length;
  return {n: Number(m[2]), at: lead + m[1].length, len: m[2].length};
}

/** Figure labels in reading order: paragraph and the place of its number. */
export function figureLabels(blocks: Block[], style: CiteStyle, kind: 'figure' | 'table' = 'figure'): Array<{para: number; at: number; len: number}> {
  const out: Array<{para: number; at: number; len: number}> = [];
  for (const b of blocks) {
    if (b.type === 'p') {
      const l = labelNumber(paragraphText(b), style, kind);
      if (l) {
        out.push({para: b.index, at: l.at, len: l.len});
      }
    }
  }
  return out;
}

const DOUBLE = 480;

/**
 * The edits that add a figure after paragraph `after`: the picture (from `path`, cx×cy EMU),
 * with a label and title when `title` is given (numbered among the figures around it; later
 * figures renumbered). Returns the edits and the figure's number (0 without a label).
 */
export function figureOps(
  blocks: Block[],
  after: ParagraphBlock,
  picture: {path: string; cx: number; cy: number; alt: string},
  title: string | null,
  style: CiteStyle,
): {ops: Op[]; n: number} {
  const labels = title !== null ? figureLabels(blocks, style) : [];
  const n = labels.filter(l => l.para <= after.index).length + 1;
  // Renumber the figures after it first: these edits keep every paragraph where it is.
  const ops: Op[] = labels
    .filter(l => l.para > after.index)
    .map((l, i) => ({op: 'text', para: l.para, start: l.at, end: l.at + l.len, text: String(n + 1 + i)}));
  const len = paragraphText(after).length;
  let p = after.index + 1;
  // A new paragraph after `after`, set out plainly: no indent, no page break, no bold/italic of its own.
  const plain = (para: number, text: string, align: 'left' | 'center'): Op[] => [
    {op: 'style', para, kind: 'normal'},
    {op: 'para', para, align, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0, pb: false},
    ...(text
      ? ([
          {op: 'text', para, start: 0, end: 0, text},
          {op: 'format', para, start: 0, end: text.length, prop: 'b', on: false},
          {op: 'format', para, start: 0, end: text.length, prop: 'i', on: false},
        ] as Op[])
      : []),
  ];
  const image = (para: number): Op => ({op: 'image', para, at: 0, path: picture.path, cx: picture.cx, cy: picture.cy, alt: picture.alt});
  ops.push({op: 'split', para: after.index, offset: len});
  const t = (title ?? '').trim();
  // Figures are centred on the page.
  if (title === null) {
    ops.push(...plain(p, '', 'center'), image(p));
    return {ops, n: 0};
  }
  if (style === 'apa') {
    const label = `Figure ${n}`;
    ops.push(...plain(p, label, 'left'), {op: 'format', para: p, start: 0, end: label.length, prop: 'b', on: true});
    ops.push({op: 'split', para: p, offset: label.length});
    p += 1;
    if (t) {
      ops.push({op: 'text', para: p, start: 0, end: 0, text: t}, {op: 'format', para: p, start: 0, end: t.length, prop: 'b', on: false}, {op: 'format', para: p, start: 0, end: t.length, prop: 'i', on: true});
    }
    ops.push({op: 'split', para: p, offset: t.length});
    p += 1;
    ops.push({op: 'para', para: p, align: 'center', first: 0}, image(p));
    return {ops, n};
  }
  const caption = style === 'mla' ? `Fig. ${n}. ${t}`.trim() : `Figure ${n}. ${t}`.trim();
  ops.push(...plain(p, '', 'center'), image(p));
  ops.push({op: 'split', para: p, offset: 1});
  p += 1;
  ops.push({op: 'para', para: p, align: 'left'});
  ops.push({op: 'text', para: p, start: 0, end: 0, text: caption}, {op: 'format', para: p, start: 0, end: caption.length, prop: 'b', on: false}, {op: 'format', para: p, start: 0, end: caption.length, prop: 'i', on: false});
  return {ops, n};
}

/**
 * The edits that add a table after paragraph `after`: rows × cols empty cells (a header row
 * when `header`), with a label and title above it when `title` is given — APA: "Table N" in
 * bold and the title in italics; MLA: "Table N" and the title; Chicago: "Table N. Title".
 * Numbered among the tables around it; later tables renumbered. A plain paragraph follows
 * the table (where the text goes on).
 */
export function tableOps(blocks: Block[], after: ParagraphBlock, size: {rows: number; cols: number; header: boolean; pct?: number}, title: string | null, style: CiteStyle): {ops: Op[]; n: number} {
  const labels = title !== null ? figureLabels(blocks, style, 'table') : [];
  const n = labels.filter(l => l.para <= after.index).length + 1;
  const ops: Op[] = labels
    .filter(l => l.para > after.index)
    .map((l, i) => ({op: 'text', para: l.para, start: l.at, end: l.at + l.len, text: String(n + 1 + i)}));
  const plain = (para: number, text: string, bold: boolean, italic: boolean): Op[] => [
    {op: 'style', para, kind: 'normal'},
    {op: 'para', para, align: 'left', line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0, pb: false},
    ...(text
      ? ([
          {op: 'text', para, start: 0, end: 0, text},
          {op: 'format', para, start: 0, end: text.length, prop: 'b', on: bold},
          {op: 'format', para, start: 0, end: text.length, prop: 'i', on: italic},
        ] as Op[])
      : []),
  ];
  let p = after.index + 1;
  ops.push({op: 'split', para: after.index, offset: paragraphText(after).length});
  const t = (title ?? '').trim();
  if (title !== null) {
    const lines: Array<{text: string; b: boolean; i: boolean}> =
      style === 'apa'
        ? [{text: `Table ${n}`, b: true, i: false}, {text: t, b: false, i: true}]
        : style === 'mla'
        ? [{text: `Table ${n}`, b: false, i: false}, {text: t, b: false, i: false}]
        : [{text: `Table ${n}. ${t}`.trim(), b: false, i: false}];
    for (const line of lines) {
      ops.push(...plain(p, line.text, line.b, line.i), {op: 'split', para: p, offset: line.text.length});
      p += 1;
    }
  }
  // The paragraph after the table: plain, no bold or italic carried over from a title.
  ops.push({op: 'para', para: p, align: 'left', first: 0, pb: false}, {op: 'tableInsert', para: -1, before: p, rows: size.rows, cols: size.cols, header: size.header, pct: size.pct ?? 100});
  return {ops, n: title !== null ? n : 0};
}
