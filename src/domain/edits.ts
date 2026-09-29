// Edits as data. The screen shows the original document with the first `cursor` edits
// applied (so Undo/Redo move the cursor), and Save sends the same list to the native
// writer (DocxEditor), which applies it to the original file. One list, two appliers:
// these functions must mean exactly what DocxEditor.format / setStyle mean.
//
// Offsets are into a paragraph's text (runs' `t` joined; objects are one U+FFFC), the
// same text the native reader and writer use.

import type {ListKind} from './lists';
import {OBJECT, paragraphText, type Block, type Comment, type Footnote, type Mark, type NotePiece, type TableCell, type PageSetup, type ParagraphBlock, type Revision, type Run, type StyleLook} from '../model/docx';

export type FormatProp = 'b' | 'i' | 'u' | 'h' | 's' | 'sup' | 'sub';
export type StyleKind = 'heading1' | 'heading2' | 'heading3' | 'title' | 'quote' | 'normal';

/** Highlight colours offered (Word's names), with what the menu calls them. */
export const HIGHLIGHTS: Array<[string, string]> = [
  ['yellow', 'Yellow'],
  ['green', 'Green'],
  ['blue', 'Blue'],
  ['magenta', 'Pink'],
  ['red', 'Red'],
  ['cyan', 'Turquoise'],
];

export type Op =
  /** value: for 'h', the highlight colour (Word's name). */
  | {op: 'format'; para: number; start: number; end: number; prop: FormatProp; on: boolean; value?: string}
  /** look: how that style looks in this document (the screen's copy; the writer ignores it). */
  | {op: 'style'; para: number; kind: StyleKind; look?: StyleLook}
  /** Replace [start, end) with text: delete when text is '', insert when start === end. */
  | {op: 'text'; para: number; start: number; end: number; text: string}
  /** Enter: the text from `offset` on becomes a new paragraph after `para`. */
  | {op: 'split'; para: number; offset: number}
  /** Backspace at the start of `para`: it joins onto the end of `para` - 1. */
  | {op: 'join'; para: number}
  /** In or out of a list. listId: a document list (number as text) or a new list's name ("n1", "b2"). */
  | {op: 'list'; para: number; kind: ListKind | 'none'; listId: string}
  /** One list level in (+1) or out (−1): Tab / Shift+Tab (DocxEditor.listLevel). */
  | {op: 'listLevel'; para: number; delta: number}
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
  | {op: 'defaults'; para: -1; font?: string; size?: number}
  /**
   * The header or footer on every page: one line aligned `align`, the text, then the page
   * number when `pageNumber`. A logo or table already there stays; para is -1.
   */
  | {op: 'headerFooter'; para: -1; kind: 'header' | 'footer'; text: string; pageNumber: boolean; align: 'left' | 'center' | 'right'}
  /** Make [start, end) a link to url. */
  | {op: 'link'; para: number; start: number; end: number; url: string}
  /** Remove the links overlapping [start, end) (a caret: the link it is in); the text stays. */
  | {op: 'unlink'; para: number; start: number; end: number}
  /**
   * Accept or reject tracked change `id` of `para` (para -1 and id '*': all of them).
   * `result`: the changed paragraphs as the file will have them (from DocxModule.preview),
   * so the screen shows exactly what the save writes; the native writer ignores it.
   */
  | {op: 'revision'; para: number; id: string; accept: boolean; result: Record<number, {runs: Run[]; revs?: Revision[]}>}
  /**
   * A new comment `id` on the text from (fromPara, from) to (toPara, to); with `parent`, a
   * reply sharing the parent's range. para is -1.
   */
  | {
      op: 'comment';
      para: -1;
      id: number;
      fromPara: number;
      from: number;
      toPara: number;
      to: number;
      text: string;
      author: string;
      initials: string;
      date: string;
      parent?: number;
    }
  /** Delete these comments (a thread: the comment and its replies). para is -1. */
  | {op: 'uncomment'; para: -1; ids: number[]}
  /**
   * A handwritten note `id` in the right margin beside character `at` of `para`: the
   * picture `png` (width×height px). Its anchor is one object character in the text.
   */
  | {op: 'ink'; para: number; at: number; id: string; png: string; width: number; height: number}
  /** Table edits (DocxEditor.table): `table` is the table's ordinal among the body's tables. */
  | {op: 'tableCell'; para: -1; table: number; row: number; cell: number; pieces: NotePiece[]}
  | {op: 'tableRowAdd'; para: -1; table: number; row: number; below: boolean}
  | {op: 'tableRowDelete'; para: -1; table: number; row: number}
  /** A new rows × cols table before paragraph `before` (the first row a header row when `header`). */
  | {op: 'tableInsert'; para: -1; before: number; rows: number; cols: number; header: boolean; pct?: number}
  /** A hidden bookmark `name` around the words a handwritten note or a quote is about (DocxEditor.addBookmark). */
  | {op: 'bookmark'; para: -1; fromPara: number; from: number; toPara: number; to: number; name: string}
  /** Removes table `table`. */
  | {op: 'tableDelete'; para: -1; table: number}
  /** Table `table` becomes `pct`% of the text width, centred. */
  | {op: 'tableWidth'; para: -1; table: number; pct: number}
  /** The picture at `at` becomes cx × cy EMU / is removed. */
  | {op: 'imageSize'; para: number; at: number; cx: number; cy: number}
  | {op: 'imageDelete'; para: number; at: number}
  /** Footnote `id` with the text `pieces`, its number at `at` (DocxEditor.addFootnote). */
  | {op: 'footnote'; para: number; at: number; id: number; pieces: NotePiece[]}
  /** Footnote `id`'s text becomes `pieces`. */
  | {op: 'footnoteSet'; para: -1; id: number; pieces: NotePiece[]}
  /** Removes footnote `id` and its number. */
  | {op: 'footnoteDelete'; para: -1; id: number}
  /** A picture from the file `path`, cx × cy EMU, in the text at `at` (DocxEditor.addImage). */
  | {op: 'image'; para: number; at: number; path: string; cx: number; cy: number; alt?: string}
  /** Remove handwritten note `id` from the document (its anchor and picture). */
  | {op: 'inkDelete'; para: number; id: string}
  /** Redefine paragraph styles (a paper format's headings); para is -1. */
  | {op: 'styleDefs'; para: -1; defs: StyleDef[]};

/** A style's new look (DocxEditor.StyleDef); missing = unchanged. */
export type StyleDef = {kind: StyleKind; font?: string; size?: number; bold?: boolean; italic?: boolean; align?: 'left' | 'center' | 'right' | 'justify'; line?: number};

/** Each style's look after the edits: the document's, with redefinitions applied. */
export function looksAfter(looks: Record<string, StyleLook> | undefined, ops: Op[]): Record<string, StyleLook> {
  const out: Record<string, StyleLook> = {...(looks ?? {})};
  for (const op of ops) {
    if (op.op === 'styleDefs') {
      for (const d of op.defs) {
        const was = out[d.kind] ?? {align: 'left', indent: 0};
        out[d.kind] = {
          ...was,
          ...(d.font !== undefined ? {bf: d.font} : {}),
          ...(d.size !== undefined ? {bs: d.size} : {}),
          ...(d.bold !== undefined ? {sb: d.bold} : {}),
          ...(d.italic !== undefined ? {si: d.italic} : {}),
          ...(d.align ? {align: d.align} : {}),
          indent: 0,
        };
      }
    }
  }
  return out;
}

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
    if (!(start >= op.start && offset <= op.end && r.t.length > 0)) {
      return r;
    }
    if (op.prop === 'h') {
      return {...r, h: op.on, hc: op.on ? op.value ?? 'yellow' : undefined};
    }
    // Superscript and subscript share one setting in Word: turning one on turns the other off.
    if (op.prop === 'sup' || op.prop === 'sub') {
      return {...r, sup: op.prop === 'sup' && op.on, sub: op.prop === 'sub' && op.on};
    }
    return {...r, [op.prop]: op.on};
  });
  return {...p, runs};
}

function styleParagraph(p: ParagraphBlock, kind: StyleKind): ParagraphBlock {
  switch (kind) {
    case 'heading1':
      return {...p, kind: 'heading', level: 1};
    case 'heading2':
      return {...p, kind: 'heading', level: 2, quote: undefined};
    case 'heading3':
      return {...p, kind: 'heading', level: 3, quote: undefined};
    case 'title':
      return {...p, kind: 'title', level: 0, quote: undefined};
    case 'quote':
      return {...p, kind: 'body', level: 0, quote: true};
    default:
      return {...p, kind: 'body', level: 0, quote: undefined};
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
    // Typed text is text: never an object, field or page break, whatever it sits beside.
    const {t: _t, obj: _o, k: _k, pg: _g, src: _s, cx: _x, cy: _y, ...format} = source ?? {t: ''};
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
  const dels = (p.revs ?? []).filter(v => v.kind === 'del');
  const first: ParagraphBlock = withRevs({...p, runs: before, sect: undefined}, [
    ...(p.revs ?? []).filter(v => v.kind === 'ins'),
    ...dels.filter(v => (v.at ?? 0) <= op.offset),
  ]);
  // Only the first half starts a new page (DocxEditor.splitParagraph).
  let second: ParagraphBlock = withRevs({...p, index: p.index + 1, runs: after, pb: undefined}, [
    ...(p.revs ?? []).filter(v => v.kind === 'ins'),
    ...dels.filter(v => (v.at ?? 0) > op.offset).map(v => ({...v, at: (v.at ?? 0) - op.offset})),
  ]);
  const marks = p.marks ?? [];
  const firstMarks = marks.filter(m => m.at < op.offset || (m.at === op.offset && m.kind !== 'start'));
  const secondMarks = marks.filter(m => !firstMarks.includes(m)).map(m => ({...m, at: m.at - op.offset}));
  (first as ParagraphBlock).marks = firstMarks.length ? firstMarks : undefined;
  second = withMarks(second, secondMarks);
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
  const shift = paragraphText(prev).length;
  out.splice(
    i - 1,
    2,
    withMarks(
      withRevs({...prev, runs: [...prev.runs, ...p.runs], sect: prev.sect || p.sect}, [
        ...(prev.revs ?? []),
        ...(p.revs ?? []).map(v => (v.kind === 'del' ? {...v, at: (v.at ?? 0) + shift} : v)),
      ]),
      [...(prev.marks ?? []), ...(p.marks ?? []).map(m => ({...m, at: m.at + shift}))],
    ),
  );
  renumber(out, i, -1);
}

/**
 * Mirrors DocxEditor.setListItem. The label is left for recount (domain/lists); a new list
 * item takes the list's first-level indent, and leaving a list drops the list indent.
 */
/** Mirrors DocxEditor.listLevel: the level changes (0–8), and the text moves a level's indent (0.5″). */
function listLevel(p: ParagraphBlock, op: Extract<Op, {op: 'listLevel'}>): ParagraphBlock {
  if (!p.num) {
    return p;
  }
  const lvl = Math.max(0, Math.min(8, p.num.lvl + op.delta));
  return {...p, num: {...p.num, lvl}, indent: Math.max(0, p.indent + 720 * (lvl - p.num.lvl))};
}

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
    if (op.op === 'headerFooter' || op.op === 'page' || op.op === 'footnoteSet') {
      continue;
    }
    if (op.op === 'tableCell' || op.op === 'tableRowAdd' || op.op === 'tableRowDelete' || op.op === 'tableInsert' || op.op === 'tableWidth' || op.op === 'tableDelete') {
      tableEdit(out, op);
      continue;
    }
    if (op.op === 'footnoteDelete') {
      const id = String(op.id);
      for (let j = 0; j < out.length; j++) {
        const b = out[j];
        if (b.type !== 'p' || !b.runs.some(r => r.fn === id)) {
          continue;
        }
        let p = b;
        for (;;) {
          let offset = 0;
          let at = -1;
          for (const r of p.runs) {
            if (at < 0 && r.obj === 'note' && r.fn === id) {
              at = offset;
            }
            offset += r.t.length;
          }
          if (at < 0) {
            break;
          }
          const runs = [...p.runs];
          runs.splice(p.runs.findIndex(r => r.obj === 'note' && r.fn === id), 1);
          p = withMarks(withRevs({...p, runs}, shiftDeletions(p.revs, at, at + 1, 0)), shiftMarks(p.marks, at, at + 1, 0));
        }
        out[j] = p;
      }
      continue;
    }
    if (op.op === 'styleDefs') {
      // Paragraphs already in a redefined style take its new look (what they set themselves stays).
      for (const d of op.defs) {
        const level = d.kind === 'heading1' ? 1 : d.kind === 'heading2' ? 2 : d.kind === 'heading3' ? 3 : 0;
        for (let j = 0; j < out.length; j++) {
          const b = out[j];
          if (b.type !== 'p' || !((level && b.kind === 'heading' && b.level === level) || (d.kind === 'title' && b.kind === 'title'))) {
            continue;
          }
          out[j] = {
            ...b,
            ...(d.font !== undefined ? {bf: d.font} : {}),
            ...(d.size !== undefined ? {bs: d.size} : {}),
            ...(d.bold !== undefined ? {sb: d.bold} : {}),
            ...(d.italic !== undefined ? {si: d.italic} : {}),
            ...(d.align && !b.ja ? {align: d.align} : {}),
            ...(!b.ji && !b.num ? {indent: 0} : {}),
          };
        }
      }
      continue;
    }
    if (op.op === 'comment' || op.op === 'uncomment') {
      commentMarks(out, op);
      continue;
    }
    if (op.op === 'bookmark') {
      const add = (index: number, m: Mark) => {
        const j = out.findIndex(b => b.type === 'p' && b.index === index);
        if (j >= 0) {
          const p = out[j] as ParagraphBlock;
          out[j] = withMarks(p, [...(p.marks ?? []), m]);
        }
      };
      add(op.fromPara, {id: op.name, kind: 'start', at: op.from});
      add(op.toPara, {id: op.name, kind: 'end', at: op.to});
      continue;
    }
    if (op.op === 'revision') {
      for (const [key, r] of Object.entries(op.result)) {
        const j = out.findIndex(b => b.type === 'p' && b.index === Number(key));
        if (j >= 0) {
          out[j] = {...(out[j] as ParagraphBlock), runs: r.runs, revs: r.revs && r.revs.length ? r.revs : undefined};
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
      case 'style': {
        // The new style's font, size, bold and italic are the file's to know: take them from
        // a paragraph already in that style, else show the kind's defaults.
        const next = styleParagraph(p, op.kind);
        if (op.look) {
          // The style as the document defines it; alignment and indent the paragraph sets
          // itself stay (as in Word), and a list item keeps its list's indent.
          const l = op.look;
          out[i] = {
            ...next,
            bf: l.bf,
            bs: l.bs,
            sb: l.sb,
            si: l.si,
            align: p.ja ? p.align : l.align,
            indent: p.ji || p.num ? p.indent : l.indent,
          };
          break;
        }
        const like = out.find(
          b => b.type === 'p' && b !== p && b.kind === next.kind && b.level === next.level && !!b.quote === !!next.quote && !b.num,
        ) as ParagraphBlock | undefined;
        out[i] = {...next, bf: like?.bf, bs: like?.bs, sb: like?.sb, si: like?.si};
        break;
      }
      case 'text':
        out[i] = withMarks(withRevs(editText(p, op), shiftDeletions(p.revs, op.start, op.end, op.text.length)), shiftMarks(p.marks, op.start, op.end, op.text.length));
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
      case 'listLevel':
        out[i] = listLevel(p, op);
        break;
      case 'para':
        out[i] = paraProps(p, op);
        break;
      case 'link':
        out[i] = {...p, runs: setLink(p.runs, op.start, op.end, true)};
        break;
      case 'imageSize':
        out[i] = {
          ...p,
          runs: splitRuns(p.runs, [op.at, op.at + 1]).map((r, k, all) => {
            const start = all.slice(0, k).reduce((n, x) => n + x.t.length, 0);
            return start === op.at && r.obj === 'image' ? {...r, cx: op.cx, cy: op.cy} : r;
          }),
        };
        break;
      case 'imageDelete': {
        let offset = 0;
        const runs: Run[] = [];
        for (const r of splitRuns(p.runs, [op.at, op.at + 1])) {
          if (!(offset === op.at && r.obj === 'image')) {
            runs.push(r);
          }
          offset += r.t.length;
        }
        out[i] = withMarks(withRevs({...p, runs}, shiftDeletions(p.revs, op.at, op.at + 1, 0)), shiftMarks(p.marks, op.at, op.at + 1, 0));
        break;
      }
      case 'footnote':
      case 'image': {
        const pic: Run = op.op === 'image' ? {t: OBJECT, obj: 'image', src: op.path, cx: op.cx, cy: op.cy} : {t: OBJECT, obj: 'note', sup: true, fn: String(op.id)};
        let offset = 0;
        const runs: Run[] = [];
        let placed = false;
        for (const r of splitRuns(p.runs, [op.at])) {
          if (!placed && offset >= op.at) {
            runs.push(pic);
            placed = true;
          }
          runs.push(r);
          offset += r.t.length;
        }
        if (!placed) {
          runs.push(pic);
        }
        out[i] = withMarks(withRevs({...p, runs}, shiftDeletions(p.revs, op.at, op.at, 1)), shiftMarks(p.marks, op.at, op.at, 1));
        break;
      }
      case 'ink': {
        let offset = 0;
        const runs: Run[] = [];
        let placed = false;
        for (const r of splitRuns(p.runs, [op.at])) {
          if (!placed && offset >= op.at) {
            runs.push({t: OBJECT, obj: 'ink', ink: op.id});
            placed = true;
          }
          runs.push(r);
          offset += r.t.length;
        }
        if (!placed) {
          runs.push({t: OBJECT, obj: 'ink', ink: op.id});
        }
        out[i] = withMarks(withRevs({...p, runs}, shiftDeletions(p.revs, op.at, op.at, 1)), shiftMarks(p.marks, op.at, op.at, 1));
        break;
      }
      case 'inkDelete': {
        let offset = 0;
        let at = -1;
        for (const r of p.runs) {
          if (r.obj === 'ink' && r.ink === op.id) {
            at = offset;
          }
          offset += r.t.length;
        }
        if (at >= 0) {
          out[i] = withMarks(
            withRevs({...p, runs: p.runs.filter(r => !(r.obj === 'ink' && r.ink === op.id))}, shiftDeletions(p.revs, at, at + 1, 0)),
            shiftMarks(p.marks, at, at + 1, 0),
          );
        }
        break;
      }
      case 'unlink': {
        const span = linkSpan(p, op.start, op.end);
        if (span) {
          out[i] = {...p, runs: setLink(p.runs, span.start, span.end, false)};
        }
        break;
      }
    }
  }
  return out;
}

function withMarks(p: ParagraphBlock, marks: Mark[] | undefined): ParagraphBlock {
  return {...p, marks: marks && marks.length ? marks : undefined};
}

/**
 * Comment anchors after [start, end) is replaced by `inserted` characters. Screen-only; as
 * the writer does, typed text joins the run before it, so marks there move after it.
 */
function shiftMarks(marks: Mark[] | undefined, start: number, end: number, inserted: number): Mark[] | undefined {
  return marks?.map(m => {
    if (m.at < start || (m.at === 0 && start === 0)) {
      return m;
    }
    return m.at >= end ? {...m, at: m.at + inserted - (end - start)} : {...m, at: start};
  });
}

/** Adds or removes comment anchors on screen, as DocxEditor.addComment / deleteComments place them. */
function commentMarks(out: Block[], op: Extract<Op, {op: 'comment' | 'uncomment'}>): void {
  const update = (index: number, change: (marks: Mark[]) => Mark[]) => {
    const j = out.findIndex(b => b.type === 'p' && b.index === index);
    if (j >= 0) {
      const p = out[j] as ParagraphBlock;
      out[j] = withMarks(p, change(p.marks ?? []));
    }
  };
  if (op.op === 'uncomment') {
    const ids = new Set(op.ids.map(String));
    for (const b of out) {
      if (b.type === 'p' && b.marks?.some(m => ids.has(m.id))) {
        update(b.index, marks => marks.filter(m => !ids.has(m.id)));
      }
    }
    return;
  }
  const id = String(op.id);
  if (op.parent !== undefined) {
    // A reply sits beside its parent's marks.
    const parent = String(op.parent);
    for (const b of out) {
      if (b.type === 'p' && b.marks?.some(m => m.id === parent)) {
        update(b.index, marks => [...marks, ...marks.filter(m => m.id === parent).map(m => ({...m, id}))]);
      }
    }
    return;
  }
  update(op.fromPara, marks => [...marks, {id, kind: 'start', at: op.from}]);
  update(op.toPara, marks => [...marks, {id, kind: 'end', at: op.to}, {id, kind: 'ref', at: op.to}]);
}

/** The comments after the edits: added ones appended, deleted ones gone. */
export function commentsAfter(comments: Comment[] | undefined, ops: Op[]): Comment[] {
  let out = comments ?? [];
  for (const op of ops) {
    if (op.op === 'comment') {
      out = [
        ...out,
        {id: String(op.id), author: op.author, initials: op.initials, date: op.date, text: op.text, parent: op.parent === undefined ? undefined : String(op.parent)},
      ];
    } else if (op.op === 'uncomment') {
      const ids = new Set(op.ids.map(String));
      out = out.filter(c => !ids.has(c.id));
    }
  }
  return out;
}

/** Every handwritten note's picture after the edits: the document's, plus those added. */
export function inksAfter(inks: Record<string, string> | undefined, ops: Op[]): Record<string, string> {
  const out = {...(inks ?? {})};
  for (const op of ops) {
    if (op.op === 'ink') {
      out[op.id] = op.png;
    }
  }
  return out;
}

/** The next free comment id. */
export function nextCommentId(comments: Comment[]): number {
  return comments.reduce((n, c) => Math.max(n, Number(c.id) + 1 || n), 0);
}

/** A comment and all its replies (replies of replies too), for deleting a thread. */
export function threadIds(comments: Comment[], id: string): number[] {
  const out = [id];
  for (let i = 0; i < out.length; i++) {
    out.push(...comments.filter(c => c.parent === out[i]).map(c => c.id));
  }
  return out.map(Number);
}

/** Mirrors DocxEditor.table: a cell's text, a row added or removed, a table inserted. */
function tableEdit(out: Block[], op: Extract<Op, {op: 'tableCell' | 'tableRowAdd' | 'tableRowDelete' | 'tableInsert' | 'tableWidth' | 'tableDelete'}>): void {
  if (op.op === 'tableDelete') {
    const k = out.findIndex(x => x.type === 'table' && x.t === op.table);
    if (k < 0) {
      return;
    }
    out.splice(k, 1);
    for (let j = k; j < out.length; j++) {
      const b = out[j];
      if (b.type === 'table') {
        out[j] = {...b, t: (b.t ?? 1) - 1};
      }
    }
    return;
  }
  if (op.op === 'tableInsert') {
    const at = out.findIndex(b => b.type === 'p' && b.index === op.before);
    if (at < 0) {
      return;
    }
    const ordinal = out.slice(0, at).filter(b => b.type === 'table').length;
    for (let j = at; j < out.length; j++) {
      const b = out[j];
      if (b.type === 'table') {
        out[j] = {...b, t: (b.t ?? 0) + 1};
      }
    }
    const grid = Array.from({length: op.rows}, () => Array.from({length: op.cols}, () => ({p: []}) as TableCell));
    const pct = op.pct ?? 100;
    out.splice(at, 0, {type: 'table', rows: op.rows, cols: op.cols, preview: '', t: ordinal, widths: Array(op.cols).fill(1), grid, ...(pct < 100 ? {wf: pct / 100, ta: 'center'} : {})});
    return;
  }
  if (op.op === 'tableWidth') {
    const k = out.findIndex(x => x.type === 'table' && x.t === op.table);
    const tb = out[k];
    if (k >= 0 && tb.type === 'table') {
      out[k] = {...tb, wf: Math.min(1, Math.max(0.2, op.pct / 100)), ta: 'center'};
    }
    return;
  }
  const j = out.findIndex(b => b.type === 'table' && b.t === op.table);
  const b = out[j];
  if (j < 0 || b.type !== 'table' || !b.grid) {
    return;
  }
  const grid = b.grid.map(r => [...r]);
  if (op.op === 'tableCell') {
    if (grid[op.row]?.[op.cell]) {
      grid[op.row][op.cell] = {...grid[op.row][op.cell], p: op.pieces.filter(x => x.t)};
    }
  } else if (op.op === 'tableRowAdd') {
    const like = grid[op.row];
    if (like) {
      grid.splice(op.below ? op.row + 1 : op.row, 0, like.map(c => ({...(c.s ? {s: c.s} : {}), p: []})));
    }
  } else if (grid.length > 1) {
    grid.splice(op.row, 1);
  }
  const preview = grid.flat().map(c => c.p.map(x => x.t).join('')).join(' ').replace(/\s+/g, ' ').trim().slice(0, 120);
  out[j] = {...b, grid, rows: grid.length, preview};
}

/** The footnotes after `ops`: added, retyped and removed ones. */
export function footnotesAfter(notes: Footnote[] | undefined, ops: Op[]): Footnote[] {
  let out = notes ?? [];
  for (const op of ops) {
    if (op.op === 'footnote') {
      out = [...out, {id: String(op.id), pieces: op.pieces}];
    } else if (op.op === 'footnoteSet') {
      out = out.map(f => (f.id === String(op.id) ? {...f, pieces: op.pieces} : f));
    } else if (op.op === 'footnoteDelete') {
      out = out.filter(f => f.id !== String(op.id));
    }
  }
  return out;
}

/** An unused footnote id (Word's are small numbers; 0 and -1 are its separators). */
export function nextFootnoteId(notes: Footnote[], blocks: Block[]): number {
  let max = 0;
  for (const f of notes) {
    max = Math.max(max, Number(f.id) || 0);
  }
  for (const b of blocks) {
    if (b.type === 'p') {
      for (const r of b.runs) {
        if (r.fn && !r.fn.startsWith('e')) {
          max = Math.max(max, Number(r.fn) || 0);
        }
      }
    }
  }
  return max + 1;
}

/**
 * Footnote numbers as Word shows them: 1, 2, 3 … in reading order (endnotes i, ii … are
 * left as they are). A paragraph is replaced only when one of its numbers changes.
 */
export function numberNotes(blocks: Block[]): Block[] {
  let n = 0;
  let e = 0;
  const roman = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
  return blocks.map(b => {
    if (b.type !== 'p' || !b.runs.some(r => r.obj === 'note')) {
      return b;
    }
    let changed = false;
    const runs = b.runs.map(r => {
      if (r.obj !== 'note') {
        return r;
      }
      const label = r.fn?.startsWith('e') ? roman[e++] ?? String(e) : String(++n);
      if (r.nn === label) {
        return r;
      }
      changed = true;
      return {...r, nn: label};
    });
    return changed ? {...b, runs} : b;
  });
}

/** The paragraph with `revs`, keeping only insertions that still have text and dropping an empty list. */
function withRevs(p: ParagraphBlock, revs: Revision[] | undefined): ParagraphBlock {
  const ids = new Set(p.runs.map(r => r.rv).filter(Boolean));
  const kept = (revs ?? []).filter(v => v.kind === 'del' || ids.has(v.id));
  return {...p, revs: kept.length ? kept : undefined};
}

/**
 * Where deletions sit after [start, end) is replaced by `inserted` characters. Screen-only
 * (accept/reject asks the file for its result), so ties go the natural way: text typed
 * where a deletion sits goes after it.
 */
function shiftDeletions(revs: Revision[] | undefined, start: number, end: number, inserted: number): Revision[] | undefined {
  return revs?.map(v => {
    if (v.kind !== 'del' || v.at === undefined) {
      return v;
    }
    if (v.at >= end && v.at > start) {
      return {...v, at: v.at + inserted - (end - start)};
    }
    return v.at > start ? {...v, at: start} : v;
  });
}

// ---------------------------------------------------------------- screen offsets

/**
 * Characters on screen that are not the paragraph's text, as {at, len, lead}: the
 * first-line indent spacer (lead, before everything) and deleted text shown struck through
 * before character `at`. The TextView's offsets count them; the model's do not.
 */
export function shownExtras(p: ParagraphBlock): Array<{at: number; len: number; lead?: boolean; caret?: boolean}> {
  const out: Array<{at: number; len: number; lead?: boolean; caret?: boolean}> = [];
  if ((p.first ?? 0) > 0 && p.runs.length > 0) {
    out.push({at: 0, len: 1, lead: true});
  }
  for (const v of p.revs ?? []) {
    if (v.kind === 'del' && v.runs?.length) {
      out.push({at: Math.max(0, v.at ?? 0), len: v.runs.reduce((n, r) => n + r.t.length, 0)});
    }
  }
  if (p.caret !== undefined) {
    // Drawn in the text while typing: after deleted text at the same place, before the character.
    out.push({at: p.caret, len: 1, caret: true});
  }
  return out;
}

/**
 * Text offsets of the paragraph's Word page breaks (the "\n" of each). On screen they are
 * plain line breaks: screen pages are not Word's pages, and text reflows by font and size
 * (CT: honouring them doubled a 50-page paper to 101 half-empty pages). Backspace at the
 * start of the line after one still removes it.
 */
export function pageBreakChars(p: ParagraphBlock): number[] {
  const out: number[] = [];
  let off = 0;
  for (const r of p.runs) {
    if (r.pg) {
      for (let i = 0; i < r.t.length; i++) {
        if (r.t[i] === '\n') {
          out.push(off + i);
        }
      }
    }
    off += r.t.length;
  }
  return out;
}


/** A text offset as a TextView offset: extras before it are counted (a deletion at the caret is after it). */
export function toShown(p: ParagraphBlock, offset: number): number {
  let n = offset;
  for (const x of shownExtras(p)) {
    if (x.lead || x.at < offset) {
      n += x.len;
    }
  }
  return n;
}

/** A TextView offset as a text offset; inside shown deleted text it snaps to where the deletion sits. */
export function fromShown(p: ParagraphBlock, shown: number): number {
  let extra = 0;
  // Extras in display order: the lead first, then deletions by position.
  const xs = shownExtras(p).sort((a, b) => (a.lead ? -1 : b.lead ? 1 : a.at - b.at || (a.caret ? 1 : 0) - (b.caret ? 1 : 0)));
  for (const x of xs) {
    const startShown = (x.lead ? 0 : x.at) + extra;
    if (shown < startShown) {
      break;
    }
    if (shown < startShown + x.len) {
      return x.lead ? 0 : x.at;
    }
    extra += x.len;
  }
  return Math.max(0, shown - extra);
}

function setLink(runs: Run[], start: number, end: number, on: boolean): Run[] {
  let offset = 0;
  return splitRuns(runs, [start, end]).map(r => {
    const s = offset;
    offset += r.t.length;
    return s >= start && offset <= end && r.t.length > 0 ? {...r, l: on} : r;
  });
}

/**
 * The whole extent of the links overlapping [start, end) — a caret (start === end) counts
 * when it touches a link — or null when there are none.
 */
export function linkSpan(p: ParagraphBlock, start: number, end: number): {start: number; end: number} | null {
  const hi = Math.max(end, start + 1);
  let pos = 0;
  let span: {start: number; end: number} | null = null;
  const runs = p.runs.map(r => {
    const s = pos;
    pos += r.t.length;
    return {r, s, e: pos};
  });
  for (const {r, s, e} of runs) {
    if (r.l && s < hi && e > (start === end ? start - 1 : start)) {
      span = span ? {start: Math.min(span.start, s), end: Math.max(span.end, e)} : {start: s, end: e};
    }
  }
  if (!span) {
    return null;
  }
  // Grow over neighbouring link runs: one link is often several runs.
  for (let grew = true; grew; ) {
    grew = false;
    for (const {r, s, e} of runs) {
      if (r.l && r.t.length > 0 && ((e === span.start && s < span.start) || (s === span.end && e > span.end))) {
        span = {start: Math.min(span.start, s), end: Math.max(span.end, e)};
        grew = true;
      }
    }
  }
  return span;
}

/** Why [start, end) of `p` can't become a link (empty, already linked, a field or object), or null. */
export function linkProblem(p: ParagraphBlock, start: number, end: number): string | null {
  if (end <= start) {
    return 'Select the words to link first.';
  }
  let pos = 0;
  for (const r of p.runs) {
    const s = pos;
    pos += r.t.length;
    if (r.t.length === 0 || pos <= start || s >= end) {
      continue;
    }
    if (r.l) {
      return 'Part of that is already a link. Remove it first.';
    }
    if (r.k || r.obj) {
      return 'A link can\'t cover a field, note or picture.';
    }
  }
  return null;
}

/**
 * What was typed as a link address, as a web address: a DOI (10.xxxx/…, doi:…) becomes
 * https://doi.org/…, www.… and bare domains get https://. Null when it doesn't look like one.
 */
export function linkUrl(typed: string): string | null {
  const t = typed.trim().replace(/[.,;]+$/, '');
  if (!t || /\s/.test(t)) {
    return null;
  }
  const doi = /^(?:doi:\s*|https?:\/\/(?:dx\.)?doi\.org\/)?(10\.\d{4,9}\/\S+)$/i.exec(t);
  if (doi) {
    return `https://doi.org/${doi[1]}`;
  }
  if (/^(https?:\/\/|mailto:)/i.test(t)) {
    return t;
  }
  if (/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(t)) {
    return `mailto:${t}`;
  }
  if (/^(www\.)?[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}(\/\S*)?$/i.test(t)) {
    return `https://${t}`;
  }
  return null;
}

export type HeaderFooter = {text: string; pageNumber: boolean; align: 'left' | 'center' | 'right' | 'justify'; other?: boolean};

/** The header or footer after the edits (the last headerFooter op of that kind wins). */
export function headerFooterAfter(current: HeaderFooter | undefined, kind: 'header' | 'footer', ops: Op[]): HeaderFooter | undefined {
  let out = current;
  for (const op of ops) {
    if (op.op === 'headerFooter' && op.kind === kind) {
      out = {text: op.text, pageNumber: op.pageNumber, align: op.align, other: current?.other};
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
      // Bold / italic may come from the paragraph's style when the run doesn't say.
      const has = prop === 'b' ? run.b ?? !!p.sb : prop === 'i' ? run.i ?? !!p.si : !!run[prop];
      if (!has) {
        return false;
      }
    }
  }
  return any;
}

/** Highlight in `colour` (null: remove highlighting) for the ranges. */
export function highlightOps(ranges: Range[], colour: string | null): Op[] {
  return ranges.map(r => ({op: 'format', para: r.para, start: r.start, end: r.end, prop: 'h', on: colour !== null, ...(colour ? {value: colour} : {})}));
}

export type Match = {para: number; start: number; end: number};

/** Every place `query` occurs in the paragraphs' text, in reading order (case-insensitive unless `matchCase`). */
export function findMatches(blocks: Block[], query: string, matchCase: boolean): Match[] {
  if (!query) {
    return [];
  }
  const q = matchCase ? query : query.toLowerCase();
  const out: Match[] = [];
  for (const b of blocks) {
    if (b.type !== 'p') {
      continue;
    }
    const text = matchCase ? paragraphText(b) : paragraphText(b).toLowerCase();
    for (let i = text.indexOf(q); i >= 0; i = text.indexOf(q, i + q.length)) {
      out.push({para: b.index, start: i, end: i + q.length});
    }
  }
  return out;
}

/**
 * Replace every match with `replacement`, as one set of edits: later matches in a
 * paragraph first, so earlier offsets stay right. Matches inside fields, form controls or
 * over objects are left alone and counted as skipped.
 */
export function replaceAllOps(blocks: Block[], matches: Match[], replacement: string): {ops: Op[]; skipped: number} {
  const byIndex = new Map<number, ParagraphBlock>();
  blocks.forEach(b => b.type === 'p' && byIndex.set(b.index, b));
  const ops: Op[] = [];
  let skipped = 0;
  const sorted = [...matches].sort((a, b) => a.para - b.para || b.start - a.start);
  for (const m of sorted) {
    const p = byIndex.get(m.para);
    if (!p || textEditProblem(p, m.start, m.end)) {
      skipped++;
      continue;
    }
    ops.push({op: 'text', para: m.para, start: m.start, end: m.end, text: replacement});
  }
  return {ops, skipped};
}

/** The edits a formatting button makes for a selection: on unless all of it already has it. */
export function formatOps(blocks: Block[], ranges: Range[], prop: FormatProp): Op[] {
  const on = !allHave(blocks, ranges, prop);
  return ranges.map(r => ({op: 'format', para: r.para, start: r.start, end: r.end, prop, on}));
}

export function styleOps(ranges: Range[], kind: StyleKind, look?: StyleLook): Op[] {
  return [...new Set(ranges.map(r => r.para))].map(para => ({op: 'style', para, kind, ...(look ? {look} : {})}));
}

const isSpace = (ch: string | undefined) => ch === undefined || /\s/.test(ch);

/**
 * A pen drag's selection, from the characters under its two ends (in order). A drag across
 * several words snaps to whole words; a drag that stays inside one word selects exactly the
 * characters it crossed — one digit for a superscript, say (CT: whole-word snapping made a
 * single character impossible to select).
 */
export function penSelection(
  textStart: string,
  textEnd: string,
  start: {para: number; char: number; offset: number},
  end: {para: number; char: number; offset: number},
): {from: Pos; to: Pos} | null {
  const ws = wordAround(textStart, start.char, 'right');
  const we = wordAround(textEnd, end.char, 'left');
  const sameWord = start.para === end.para && !!ws && !!we && ws.start === we.start && ws.end === we.end && !isSpace(textStart[start.char]);
  if (sameWord) {
    const lo = Math.min(start.char, end.char);
    const hi = Math.max(start.char, end.char) + 1;
    return {from: {para: start.para, offset: lo}, to: {para: start.para, offset: Math.min(hi, textStart.length)}};
  }
  const from = {para: start.para, offset: ws?.start ?? start.offset};
  const to = {para: end.para, offset: we?.end ?? end.offset};
  return comparePos(from, to) < 0 ? {from, to} : null;
}

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
