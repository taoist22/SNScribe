// The review list (View → Comments and changes…): every comment thread, handwritten note
// and tracked change in reading order, with the words each is about. A row opens its page
// and the review pane; the pane's ◀ ▶ step through the same list.

import {paragraphText, type Block, type Comment} from '../model/docx';
import {threadIds} from './edits';
import {INK_LINK} from './links';

export type ReviewKind = 'comment' | 'ink' | 'change';

export type ReviewItem = {
  kind: ReviewKind;
  /** The review pane's id for it: the thread id, INK_LINK + the note's id, or "para|change id". */
  id: string;
  /** Where it is: the block (for the page), the paragraph and the text offset. */
  block: number;
  para: number;
  offset: number;
  /** The words it is about (comments and notes), shortened. */
  words: string;
  author: string;
  /** The comment, or the inserted or deleted text. */
  text: string;
  date: string;
  replies: number;
  del?: boolean;
  move?: boolean;
};

const SHORT = 90;

function tidy(t: string): string {
  const s = t.replace(/￼/g, '').replace(/\s+/g, ' ').trim();
  return s.length > SHORT ? `${s.slice(0, SHORT - 1).trimEnd()}…` : s;
}

/** The words from a range's start (block `from`, offset `at`) to its end mark, across paragraphs. */
function rangeWords(blocks: Block[], from: number, at: number, id: string): string {
  let out = '';
  for (let bi = from; bi < blocks.length && out.length < SHORT * 2; bi++) {
    const b = blocks[bi];
    if (b.type !== 'p') {
      continue;
    }
    const t = paragraphText(b);
    const s = bi === from ? at : 0;
    const end = (b.marks ?? []).find(m => m.id === id && m.kind === 'end' && m.at >= s);
    out += (out ? ' ' : '') + t.slice(s, end ? end.at : t.length);
    if (end) {
      break;
    }
  }
  return tidy(out);
}

/** Every comment thread, handwritten note and tracked change, in reading order. */
export function reviewItems(blocks: Block[], comments: Comment[]): ReviewItem[] {
  const threads = new Map(comments.filter(c => !c.parent).map(c => [c.id, c]));
  const out: ReviewItem[] = [];
  blocks.forEach((b, bi) => {
    if (b.type !== 'p') {
      return;
    }
    const marks = b.marks ?? [];
    // Comments: at the start of their range (or, with no range, at the reference mark).
    for (const m of marks) {
      const c = threads.get(m.id);
      const hasStart = marks.some(x => x.id === m.id && x.kind === 'start');
      if (!c || (m.kind !== 'start' && !(m.kind === 'ref' && !hasStart))) {
        continue;
      }
      out.push({
        kind: 'comment',
        id: c.id,
        block: bi,
        para: b.index,
        offset: m.at,
        words: m.kind === 'start' ? rangeWords(blocks, bi, m.at, m.id) : '',
        author: c.author,
        text: tidy(c.text || (c.pictures ? '✎ handwritten note' : '')),
        date: c.date,
        replies: threadIds(comments, c.id).length - 1,
      });
    }
    let off = 0;
    const firstOf = new Map<string, number>();
    for (const r of b.runs) {
      if (r.obj === 'ink' && r.ink) {
        const name = `${INK_LINK}${r.ink}`;
        const start = marks.find(m => m.id === name && m.kind === 'start');
        out.push({kind: 'ink', id: name, block: bi, para: b.index, offset: off, words: start ? rangeWords(blocks, bi, start.at, name) : '', author: '', text: '', date: '', replies: 0});
      }
      if (r.rv && !firstOf.has(r.rv)) {
        firstOf.set(r.rv, off);
      }
      off += r.t.length;
    }
    for (const v of b.revs ?? []) {
      const del = v.kind === 'del';
      const text = del ? (v.runs ?? []).map(r => r.t).join('') : b.runs.filter(r => r.rv === v.id).map(r => r.t).join('');
      out.push({
        kind: 'change',
        id: `${b.index}|${v.id}`,
        block: bi,
        para: b.index,
        offset: del ? Math.max(0, v.at ?? 0) : firstOf.get(v.id) ?? 0,
        words: '',
        author: v.author,
        text: tidy(text),
        date: v.date,
        replies: 0,
        del,
        move: v.move,
      });
    }
  });
  const rank: Record<ReviewKind, number> = {comment: 0, ink: 1, change: 2};
  return out.sort((a, b) => a.block - b.block || a.offset - b.offset || rank[a.kind] - rank[b.kind]);
}
