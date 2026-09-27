// Edits as data. The screen shows the original document with the first `cursor` edits
// applied (so Undo/Redo move the cursor), and Save sends the same list to the native
// writer (DocxEditor), which applies it to the original file. One list, two appliers:
// these functions must mean exactly what DocxEditor.format / setStyle mean.
//
// Offsets are into a paragraph's text (runs' `t` joined; objects are one U+FFFC), the
// same text the native reader and writer use.

import {OBJECT, paragraphText, type Block, type PageSetup, type ParagraphBlock, type Run} from '../model/docx';

export type FormatProp = 'b' | 'i' | 'u' | 'h';
export type StyleKind = 'heading1' | 'heading2' | 'title' | 'normal';

export type Op =
  | {op: 'format'; para: number; start: number; end: number; prop: FormatProp; on: boolean}
  | {op: 'style'; para: number; kind: StyleKind}
  /** Replace [start, end) with text: delete when text is '', insert when start === end. */
  | {op: 'text'; para: number; start: number; end: number; text: string}
  /** Enter: the text from `offset` on becomes a new paragraph after `para`. */
  | {op: 'split'; para: number; offset: number}
  /** Backspace at the start of `para`: it joins onto the end of `para` - 1. */
  | {op: 'join'; para: number}
  /** In or out of a list. listId: a document list (number as text) or a new list's name ("n1", "b2"). */
  | {op: 'list'; para: number; kind: 'number' | 'bullet' | 'none'; listId: string}
  /** Font family and/or size (half-points) of [start, end). */
  | {op: 'runStyle'; para: number; start: number; end: number; font?: string; size?: number}
  /** Paragraph formatting; missing = unchanged. first: twips, negative = hanging, 0 = none. */
  | {
      op: 'para';
      para: number;
      align?: ParagraphBlock['align'];
      line?: number;
      lineRule?: string;
      before?: number;
      after?: number;
      first?: number;
      pb?: boolean;
    }
  /** Page size, orientation and margins (twips) for the whole document; para is -1. */
  | {op: 'page'; para: -1; width?: number; height?: number; landscape?: boolean; top?: number; right?: number; bottom?: number; left?: number}
  /** The document's default font and size (half-points), for text without its own; para is -1. */
  | {op: 'defaults'; para: -1; font?: string; size?: number};

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

/** Every paragraph's text after `ops`, in order: sent with Save as a cross-check. */
export function expectedTexts(blocks: Block[], ops: Op[]): string[] {
  return applyOps(blocks, ops)
    .filter((b): b is ParagraphBlock => b.type === 'p')
    .map(paragraphText);
}

/** Paragraph indexes after position `from` in `blocks` move by `delta` (splits and joins renumber). */
function renumber(blocks: Block[], from: number, delta: number): void {
  for (let j = from; j < blocks.length; j++) {
    const b = blocks[j];
    if (b.type === 'p') {
      blocks[j] = {...b, index: b.index + delta};
    }
  }
}

/**
 * Mirrors DocxEditor.splitParagraph: both halves keep the paragraph's properties; the
 * section break stays with the second; Enter at the end of a heading or title gives a
 * body paragraph (Word's "next" style, approximately — the file uses the real one).
 */
function split(out: Block[], i: number, op: Extract<Op, {op: 'split'}>): void {
  const p = out[i] as ParagraphBlock;
  const len = paragraphText(p).length;
  const runs = splitRuns(p.runs, [op.offset]);
  let offset = 0;
  const before: Run[] = [];
  const after: Run[] = [];
  for (const r of runs) {
    (offset < op.offset ? before : after).push(r);
    offset += r.t.length;
  }
  const first: ParagraphBlock = {...p, runs: before, sect: undefined};
  let second: ParagraphBlock = {...p, index: p.index + 1, runs: after};
  if (op.offset === len && p.kind !== 'body') {
    second = {...second, kind: 'body', level: 0, style: ''};
  }
  out.splice(i, 1, first, second);
  renumber(out, i + 2, 1);
}

/** Mirrors DocxEditor.joinParagraph: the first paragraph's properties, both texts. */
function join(out: Block[], i: number): void {
  const p = out[i] as ParagraphBlock;
  const prev = out[i - 1] as ParagraphBlock;
  out.splice(i - 1, 2, {...prev, runs: [...prev.runs, ...p.runs], sect: prev.sect || p.sect});
  renumber(out, i, -1);
}

/**
 * Mirrors DocxEditor.setListItem. The label is left for recount (domain/lists); a new list
 * item takes the list's first-level indent, and leaving a list drops the list indent.
 */
function listItem(p: ParagraphBlock, op: Extract<Op, {op: 'list'}>): ParagraphBlock {
  if (op.kind === 'none') {
    return {...p, num: undefined, list: undefined, indent: p.num ? 0 : p.indent};
  }
  const id = /^\d+$/.test(op.listId) ? Number(op.listId) : op.listId;
  // The paragraph's own indents go: a list item is indented by its list (as the writer does).
  return {...p, num: {id, lvl: 0}, list: p.list ?? '', indent: 720, first: undefined};
}

/**
 * Mirrors DocxEditor.paraProps: a hanging indent brings a left indent at least as deep,
 * and taking the hanging indent away takes that left indent back.
 */
function paraProps(p: ParagraphBlock, op: Extract<Op, {op: 'para'}>): ParagraphBlock {
  const out: ParagraphBlock = {...p};
  if (op.align) {
    out.align = op.align;
  }
  if (op.line !== undefined) {
    out.line = op.line;
    out.lineRule = op.lineRule ?? 'auto';
  }
  if (op.before !== undefined) {
    out.before = op.before;
  }
  if (op.after !== undefined) {
    out.after = op.after;
  }
  if (op.first !== undefined) {
    if (op.first < 0) {
      out.first = op.first;
      out.indent = Math.max(p.indent, -op.first);
    } else if (op.first > 0) {
      out.first = op.first;
    } else {
      if (p.first !== undefined && p.first < 0 && p.indent === -p.first) {
        out.indent = 0;
      }
      out.first = undefined;
    }
  }
  if (op.pb !== undefined) {
    out.pb = op.pb || undefined;
  }
  return out;
}

/** Mirrors DocxEditor.runStyle. */
function runStyle(p: ParagraphBlock, op: Extract<Op, {op: 'runStyle'}>): ParagraphBlock {
  if (op.end <= op.start) {
    return p;
  }
  let offset = 0;
  const runs = splitRuns(p.runs, [op.start, op.end]).map(r => {
    const s = offset;
    offset += r.t.length;
    if (r.t.length === 0 || s < op.start || offset > op.end) {
      return r;
    }
    return {...r, ...(op.font !== undefined ? {f: op.font} : {}), ...(op.size !== undefined ? {sz: op.size} : {})};
  });
  return {...p, runs};
}

/** The document with `ops` applied, in order. Untouched blocks keep their identity. */
export function applyOps(blocks: Block[], ops: Op[]): Block[] {
  if (ops.length === 0) {
    return blocks;
  }
  const out = blocks.slice();
  for (const op of ops) {
    if (op.op === 'defaults') {
      // The document default is the base of every paragraph without a style font/size of its own;
      // on screen, apply it to all (styles that set their own are rare in papers).
      for (let j = 0; j < out.length; j++) {
        const b = out[j];
        if (b.type === 'p') {
          out[j] = {...b, ...(op.font ? {bf: op.font} : {}), ...(op.size ? {bs: op.size} : {})};
        }
      }
      continue;
    }
    const i = out.findIndex(b => b.type === 'p' && b.index === op.para);
    if (i < 0) {
      continue;
    }
    const p = out[i] as ParagraphBlock;
    switch (op.op) {
      case 'format':
        out[i] = formatParagraph(p, op);
        break;
      case 'style':
        out[i] = styleParagraph(p, op.kind);
        break;
      case 'text':
        out[i] = editText(p, op);
        break;
      case 'split':
        split(out, i, op);
        break;
      case 'join':
        if (i > 0 && out[i - 1].type === 'p') {
          join(out, i);
        }
        break;
      case 'list':
        out[i] = listItem(p, op);
        break;
      case 'runStyle':
        out[i] = runStyle(p, op);
        break;
      case 'para':
        out[i] = paraProps(p, op);
        break;
    }
  }
  return out;
}

/** Why Enter can't split `p` at `offset` (inside a field or form control), or null. */
export function splitProblem(p: ParagraphBlock, offset: number): string | null {
  let pos = 0;
  let before: Run | null = null;
  let after: Run | null = null;
  for (const r of p.runs) {
    const s = pos;
    pos += r.t.length;
    if (r.t.length === 0) {
      continue;
    }
    if (r.k && offset > s && offset < pos) {
      return 'A new paragraph can\'t start inside a field or form control.';
    }
    if (pos === offset) {
      before = r;
    }
    if (s === offset && !after) {
      after = r;
    }
  }
  return before?.k && after?.k ? 'A new paragraph can\'t start inside a field or form control.' : null;
}

/** Why paragraph `para` can't be joined onto the one before it, or null. Mirrors the native refusal. */
export function joinProblem(blocks: Block[], para: number): string | null {
  const i = blocks.findIndex(b => b.type === 'p' && b.index === para);
  if (i <= 0) {
    return 'There is no paragraph before this one.';
  }
  const prev = blocks[i - 1];
  if (prev.type !== 'p') {
    return 'A table or other block sits before this paragraph, so it can\'t be joined to it.';
  }
  if (prev.sect) {
    return 'A section break (page layout change) separates these paragraphs.';
  }
  return null;
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

/** Page setup after the edits (page ops apply in order over the document's own). */
export function pageAfter(page: PageSetup | undefined, ops: Op[]): PageSetup {
  let out: PageSetup = page ?? {width: 12240, height: 15840, top: 1440, right: 1440, bottom: 1440, left: 1440, landscape: false};
  for (const op of ops) {
    if (op.op === 'page') {
      const {op: _o, para: _p, ...rest} = op;
      out = {...out, ...Object.fromEntries(Object.entries(rest).filter(([, v]) => v !== undefined))};
    }
  }
  return out;
}
