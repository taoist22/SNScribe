// What each page of the page map holds, for its card in the Pages view: the headings that
// start on it, its opening words, and how many comments, handwritten notes and tracked
// changes are on it — so pages that need attention are easy to find.

import {paragraphText, type Block} from '../model/docx';
import {fromShown} from './edits';
import type {PageStart} from './paging';

export type PageInfo = {
  /** The first heading that starts on this page ('' if none does). */
  heading: string;
  /** The page's opening words. */
  text: string;
  /** Comment threads anchored here (replies are not counted again). */
  comments: number;
  /** Handwritten margin notes anchored here. */
  notes: number;
  /** Tracked insertions and deletions here. */
  changes: number;
};

/**
 * Page `i` of `pages` over `blocks` (the text the pages were counted from). A paragraph that
 * runs across pages is split where the next page starts: what sits before that character is
 * on this page. `isReply(id)`: whether comment `id` answers another one.
 */
export function pageInfo(blocks: Block[], pages: PageStart[], i: number, isReply: (id: string) => boolean): PageInfo {
  const start = pages[i];
  const end = pages[i + 1];
  const info: PageInfo = {heading: '', text: '', comments: 0, notes: 0, changes: 0};
  if (!start) {
    return info;
  }
  // The last block with anything on this page: the next page's first block too, when that
  // page starts part-way down it.
  const last = end ? (end.anchor.offset > 0 ? end.anchor.block : end.anchor.block - 1) : blocks.length - 1;
  let text = '';
  for (let bi = start.anchor.block; bi <= last && bi < blocks.length; bi++) {
    const b = blocks[bi];
    const from = bi === start.anchor.block && b.type === 'p' && start.char > 0 ? fromShown(b, start.char) : 0;
    const to = end && bi === end.anchor.block && b.type === 'p' ? fromShown(b, end.char) : Infinity;
    if (text.length < 140) {
      const t = b.type === 'p' ? paragraphText(b) : b.type === 'table' ? '[Table]' : `[${b.what}]`;
      text += (text ? ' ' : '') + (b.type === 'p' ? t.slice(from, to === Infinity ? undefined : to) : t);
    }
    if (b.type !== 'p') {
      continue;
    }
    const here = (at: number) => at >= from && at < to;
    const startsHere = bi > start.anchor.block || start.anchor.offset === 0;
    if (!info.heading && startsHere && b.kind !== 'body') {
      info.heading = paragraphText(b).replace(/\s+/g, ' ').trim();
    }
    let off = 0;
    const inserted = new Map<string, number>();
    for (const r of b.runs) {
      if (r.obj === 'ink' && here(off)) {
        info.notes++;
      }
      if (r.rv && !inserted.has(r.rv)) {
        inserted.set(r.rv, off);
      }
      off += r.t.length;
    }
    for (const m of b.marks ?? []) {
      if (m.kind === 'ref' && here(m.at) && !isReply(m.id)) {
        info.comments++;
      }
    }
    for (const v of b.revs ?? []) {
      const at = v.kind === 'del' ? v.at ?? 0 : inserted.get(v.id);
      if (at !== undefined && here(at)) {
        info.changes++;
      }
    }
  }
  info.text = text.replace(/\s+/g, ' ').replace(/￼/g, '').trim().slice(0, 140);
  return info;
}

/** "Comments 2 · Note 1 · Changes 3", or '' for a page with none. */
export function pageMarks(p: PageInfo): string {
  const part = (n: number, one: string, many: string) => (n ? `${n === 1 ? one : many} ${n}` : '');
  return [part(p.comments, 'Comment', 'Comments'), part(p.notes, 'Note', 'Notes'), part(p.changes, 'Change', 'Changes')].filter(Boolean).join(' · ');
}
