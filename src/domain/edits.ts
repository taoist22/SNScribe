// Edits as data. The screen shows the original document with the first `cursor` edits
// applied (so Undo/Redo move the cursor), and Save sends the same list to the native
// writer (DocxEditor), which applies it to the original file. One list, two appliers:
// these functions must mean exactly what DocxEditor.format / setStyle mean.
//
// Offsets are into a paragraph's text (runs' `t` joined; objects are one U+FFFC), the
// same text the native reader and writer use.

import {OBJECT, paragraphText, type Block, type ParagraphBlock, type Run} from '../model/docx';

export type FormatProp = 'b' | 'i' | 'u' | 'h';
export type StyleKind = 'heading1' | 'heading2' | 'title' | 'normal';

export type Op =
  | {op: 'format'; para: number; start: number; end: number; prop: FormatProp; on: boolean}
  | {op: 'style'; para: number; kind: StyleKind}
  /** Replace [start, end) with text: delete when text is '', insert when start === end. */
  | {op: 'text'; para: number; start: number; end: number; text: string};

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

/** The run holding character `ch`, when it is a text run (not an object). */
function runAt(runs: Run[], ch: number): Run | null {
  if (ch < 0) {
    return null;
  }
  let offset = 0;
  for (const r of runs) {
    const end = offset + r.t.length;
    if (ch >= offset && ch < end) {
      return r.obj ? null : r;
    }
    offset = end;
  }
  return null;
}

/**
 * Mirrors DocxEditor.replaceText. New text takes the formatting of the first replaced
 * character, or when inserting, of the character before it, else the one after; it goes
 * between the run ending at `start` and the run starting there.
 */
function editText(p: ParagraphBlock, op: Extract<Op, {op: 'text'}>): ParagraphBlock {
  const source = runAt(p.runs, op.end > op.start ? op.start : op.start - 1) ?? runAt(p.runs, op.start);
  let offset = 0;
  const kept: Run[] = [];
  let insertAt = -1;
  for (const r of splitRuns(p.runs, [op.start, op.end])) {
    const s = offset;
    offset += r.t.length;
    if (r.t.length > 0 && s >= op.start && offset <= op.end) {
      continue;
    }
    if (insertAt < 0 && s >= op.start && r.t.length > 0) {
      insertAt = kept.length;
    }
    kept.push(r);
  }
  if (op.text) {
    const {t: _t, obj: _o, k: _k, ...format} = source ?? {t: ''};
    kept.splice(insertAt < 0 ? kept.length : insertAt, 0, {...format, t: op.text});
  }
  return {...p, runs: kept};
}

/**
 * Why a text edit of [start, end) can't be made, or null when it can: the range must not
 * touch an image, note marker, other object, or field / content-control text, and typing
 * must not land inside a field. Mirrors the native refusal.
 */
export function textEditProblem(p: ParagraphBlock, start: number, end: number): string | null {
  let offset = 0;
  for (const r of p.runs) {
    const s = offset;
    offset += r.t.length;
    if (r.t.length > 0 && s < end && offset > start && (r.obj || r.k || r.t.includes(OBJECT))) {
      return r.k ? 'That text belongs to a field or form control, which can\'t be edited yet.' : 'The selection includes an image or note marker, which can\'t be deleted yet.';
    }
  }
  const source = runAt(p.runs, end > start ? start : start - 1) ?? runAt(p.runs, start);
  return source?.k ? 'Typing inside a field or form control isn\'t supported yet.' : null;
}

/**
 * The range a Delete removes: the selection, plus one neighbouring space when deleting a
 * whole word would otherwise leave two spaces (or a space before punctuation).
 */
export function deletionRange(text: string, start: number, end: number): {start: number; end: number} {
  const before = text[start - 1];
  const after = text[end];
  if (after === ' ' && (start === 0 || before === ' ')) {
    return {start, end: end + 1};
  }
  if (before === ' ' && (end === text.length || after === ' ' || /[.,;:!?)]/.test(after ?? ''))) {
    return {start: start - 1, end};
  }
  return {start, end};
}

/**
 * Where a position measured before a text edit ends up after it: the edit replaced
 * [r.start, r.end) of paragraph r.para with `len` characters.
 */
export function shiftPos(p: Pos, r: Range, len: number): Pos {
  if (p.para !== r.para || p.offset <= r.start) {
    return p;
  }
  if (p.offset >= r.end) {
    return {para: p.para, offset: p.offset + len - (r.end - r.start)};
  }
  return {para: p.para, offset: r.start + len};
}

/** Paragraph texts after `ops`, for the paragraphs they touch: sent with Save as a cross-check. */
export function expectedTexts(blocks: Block[], ops: Op[]): Record<number, string> {
  const after = applyOps(blocks, ops);
  const touched = new Set(ops.map(o => o.para));
  const out: Record<number, string> = {};
  for (const b of after) {
    if (b.type === 'p' && touched.has(b.index)) {
      out[b.index] = paragraphText(b);
    }
  }
  return out;
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
    out[i] = op.op === 'format' ? formatParagraph(p, op) : op.op === 'style' ? styleParagraph(p, op.kind) : editText(p, op);
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
