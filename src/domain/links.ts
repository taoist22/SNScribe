// Words linked to something in the margin or the document's sources — the words a comment
// is on, the words a handwritten note is about, an inserted quote — shown with a dashed
// underline; a tap on them opens the comment, note or quote. Comments carry their range in
// Word; notes and quotes carry it in a hidden bookmark SNScribe adds ("_sns_ink_<id>",
// "_sns_q_<quote>_<n>").

import {paragraphText, type Block} from '../model/docx';

export type LinkKind = 'comment' | 'ink' | 'quote';

/** One stretch of linked words inside one paragraph. */
export type LinkSegment = {id: string; kind: LinkKind; start: number; end: number};

export const INK_LINK = '_sns_ink_';
export const QUOTE_LINK = '_sns_q_';

/** What a mark id links to: a comment thread, a handwritten note, a quote; else nothing. */
export function linkKind(id: string, threads: Set<string>): LinkKind | null {
  if (id.startsWith(INK_LINK)) {
    return 'ink';
  }
  if (id.startsWith(QUOTE_LINK)) {
    return 'quote';
  }
  return threads.has(id) ? 'comment' : null;
}

/** The bookmark name for a quote: its saved quote's id (shortened) and a number that makes it unique. */
export function quoteLinkName(quoteId: string): string {
  const q = quoteId.replace(/[^A-Za-z0-9]/g, '').slice(0, 14);
  return `${QUOTE_LINK}${q}_${Date.now().toString(36).slice(-6)}`;
}

/** The saved quote a quote link names (matched by the id's start). */
export function quoteOfLink(name: string, quotes: Array<{id: string}>): {id: string} | undefined {
  const q = name.slice(QUOTE_LINK.length).split('_')[0];
  return q ? quotes.find(x => x.id.replace(/[^A-Za-z0-9]/g, '').startsWith(q)) : undefined;
}

/**
 * Every paragraph's linked stretches, by paragraph index. A range may run across
 * paragraphs: it covers the rest of its first, all of the ones between, and the start of
 * its last. Empty stretches are left out.
 */
export function linkSegments(blocks: Block[], threads: Set<string>): Map<number, LinkSegment[]> {
  const out = new Map<number, LinkSegment[]>();
  const open = new Map<string, LinkKind>();
  for (const b of blocks) {
    if (b.type !== 'p') {
      continue;
    }
    const len = paragraphText(b).length;
    const segs: LinkSegment[] = [];
    const startAt = new Map<string, number>();
    for (const id of open.keys()) {
      startAt.set(id, 0);
    }
    const events = [...(b.marks ?? [])].filter(m => m.kind !== 'ref').sort((x, y) => x.at - y.at || (x.kind === 'end' ? -1 : 1));
    for (const m of events) {
      const kind = linkKind(m.id, threads);
      if (!kind) {
        continue;
      }
      if (m.kind === 'start') {
        open.set(m.id, kind);
        startAt.set(m.id, m.at);
      } else if (open.has(m.id)) {
        const s = startAt.get(m.id) ?? 0;
        if (m.at > s) {
          segs.push({id: m.id, kind: open.get(m.id)!, start: s, end: m.at});
        }
        open.delete(m.id);
        startAt.delete(m.id);
      }
    }
    for (const [id, s] of startAt) {
      if (open.has(id) && len > s) {
        segs.push({id, kind: open.get(id)!, start: s, end: len});
      }
    }
    if (segs.length) {
      out.set(b.index, segs);
    }
  }
  return out;
}
