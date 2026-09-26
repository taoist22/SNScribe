// Edits as data. The screen shows the original document with the first `cursor` edits
// applied (so Undo/Redo move the cursor), and Save sends the same list to the native
// writer (DocxEditor), which applies it to the original file. One list, two appliers:
// these functions must mean exactly what DocxEditor.format / setStyle mean.
//
// Offsets are into a paragraph's text (runs' `t` joined; objects are one U+FFFC), the
// same text the native reader and writer use.

import {paragraphText, type Block, type ParagraphBlock, type Run} from '../model/docx';

export type FormatProp = 'b' | 'i' | 'u' | 'h';
export type StyleKind = 'heading1' | 'heading2' | 'title' | 'normal';

export type Op =
  | {op: 'format'; para: number; start: number; end: number; prop: FormatProp; on: boolean}
  | {op: 'style'; para: number; kind: StyleKind};

/** A position between characters of paragraph `para` (Paragraph.index). */
export type Pos = {para: number; offset: number};
export type Range = {para: number; start: number; end: number};

// ---------------------------------------------------------------- applying edits

/** Runs cut at every offset in `cuts`. Object runs (one U+FFFC) are never cut. */
export function splitRuns(runs: Run[], cuts: number[]): Run[] {
  const out: Run[] = [];
  let offset = 0;
  for (const r of runs) {
    const end = offset + r.t.length;
    const inside = r.obj ? [] : [...new Set(cuts.filter(c => c > offset && c < end))].sort((a, b) => a - b);
    let from = offset;
    for (const c of [...inside, end]) {
      out.push({...r, t: r.t.slice(from - offset, c - offset)});
      from = c;
    }
    offset = end;
  }
  return out;
}

function formatParagraph(p: ParagraphBlock, op: Extract<Op, {op: 'format'}>): ParagraphBlock {
  if (op.end <= op.start) {
    return p;
  }
  let offset = 0;
  const runs = splitRuns(p.runs, [op.start, op.end]).map(r => {
    const start = offset;
    offset += r.t.length;
    return start >= op.start && offset <= op.end && r.t.length > 0 ? {...r, [op.prop]: op.on} : r;
  });
  return {...p, runs};
}

function styleParagraph(p: ParagraphBlock, kind: StyleKind): ParagraphBlock {
  switch (kind) {
    case 'heading1':
      return {...p, kind: 'heading', level: 1};
    case 'heading2':
      return {...p, kind: 'heading', level: 2};
    case 'title':
      return {...p, kind: 'title', level: 0};
    default:
      return {...p, kind: 'body', level: 0};
  }
}

/** The document with `ops` applied, in order. Untouched blocks keep their identity. */
export function applyOps(blocks: Block[], ops: Op[]): Block[] {
  if (ops.length === 0) {
    return blocks;
  }
  const out = blocks.slice();
  const at = new Map<number, number>();
  blocks.forEach((b, i) => b.type === 'p' && at.set(b.index, i));
  for (const op of ops) {
    const i = at.get(op.para);
    if (i === undefined) {
      continue;
    }
    const p = out[i] as ParagraphBlock;
    out[i] = op.op === 'format' ? formatParagraph(p, op) : styleParagraph(p, op.kind);
  }
  return out;
}

// ---------------------------------------------------------------- selections

export function comparePos(a: Pos, b: Pos): number {
  return a.para !== b.para ? a.para - b.para : a.offset - b.offset;
}

/** A selection as one non-empty [start, end) per paragraph it covers, in document order. */
export function rangesBetween(blocks: Block[], a: Pos, b: Pos): Range[] {
  const [from, to] = comparePos(a, b) <= 0 ? [a, b] : [b, a];
  const out: Range[] = [];
  for (const blk of blocks) {
    if (blk.type !== 'p' || blk.index < from.para || blk.index > to.para) {
      continue;
    }
    const len = paragraphText(blk).length;
    const start = blk.index === from.para ? from.offset : 0;
    const end = blk.index === to.para ? to.offset : len;
    if (end > start) {
      out.push({para: blk.index, start, end});
    }
  }
  return out;
}

/** Whether every selected character already has `prop` — then the button turns it off. */
export function allHave(blocks: Block[], ranges: Range[], prop: FormatProp): boolean {
  const byIndex = new Map<number, ParagraphBlock>();
  blocks.forEach(b => b.type === 'p' && byIndex.set(b.index, b));
  let any = false;
  for (const r of ranges) {
    const p = byIndex.get(r.para);
    if (!p) {
      continue;
    }
    let offset = 0;
    for (const run of p.runs) {
      const s = offset;
      offset += run.t.length;
      if (offset <= r.start || s >= r.end || run.t.trim() === '') {
        continue;
      }
      any = true;
      if (!run[prop]) {
        return false;
      }
    }
  }
  return any;
}

/** The edits a formatting button makes for a selection: on unless all of it already has it. */
export function formatOps(blocks: Block[], ranges: Range[], prop: FormatProp): Op[] {
  const on = !allHave(blocks, ranges, prop);
  return ranges.map(r => ({op: 'format', para: r.para, start: r.start, end: r.end, prop, on}));
}

export function styleOps(ranges: Range[], kind: StyleKind): Op[] {
  return [...new Set(ranges.map(r => r.para))].map(para => ({op: 'style', para, kind}));
}

const isSpace = (ch: string | undefined) => ch === undefined || /\s/.test(ch);

/**
 * The word containing character `ch`; on whitespace, the next word in `dir`. A selection's
 * start looks right and its end looks left, so a pen in a gap never grabs the far word.
 */
export function wordAround(text: string, ch: number, dir: 'left' | 'right'): {start: number; end: number} | null {
  let i = Math.max(0, Math.min(ch, text.length - 1));
  while (i >= 0 && i < text.length && isSpace(text[i])) {
    i += dir === 'right' ? 1 : -1;
  }
  if (i < 0 || i >= text.length) {
    return null;
  }
  let start = i;
  while (start > 0 && !isSpace(text[start - 1])) {
    start--;
  }
  let end = i + 1;
  while (end < text.length && !isSpace(text[end])) {
    end++;
  }
  return {start, end};
}

/** Runs cut at the selection, the selected ones flagged, for drawing. */
export function markSelection(runs: Run[], sel: {start: number; end: number} | null): Array<Run & {sel?: boolean}> {
  if (!sel) {
    return runs;
  }
  let offset = 0;
  return splitRuns(runs, [sel.start, sel.end]).map(r => {
    const s = offset;
    offset += r.t.length;
    return s >= sel.start && offset <= sel.end && r.t.length > 0 ? {...r, sel: true} : r;
  });
}
