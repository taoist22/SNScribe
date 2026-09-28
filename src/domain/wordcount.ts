// Word counts for a word-count target: the paper's body text, as most assignment limits
// count it — without the title page and without the reference list (and anything after
// it, such as appendices). Headings and block quotes count.

import {countWords, paragraphText, type Block, type ParagraphBlock} from '../model/docx';
import {findReferenceHeading} from './citations';

const counted = new WeakMap<Block, number>();

/** Words in one paragraph, remembered per block (edits make new blocks). */
export function paragraphWords(b: ParagraphBlock): number {
  let n = counted.get(b);
  if (n === undefined) {
    n = countWords([paragraphText(b)]).words;
    counted.set(b, n);
  }
  return n;
}

/** How far a title page reaches: a page break within the first paragraphs ends it. */
const TITLE_PAGE_MAX = 15;

/**
 * The paragraphs that make up the body: after a title page (the paragraphs before a page
 * break near the start — the paragraph's own "starts on a new page" or a Word page break)
 * and before the reference list heading.
 */
export function bodyRange(blocks: Block[]): {from: number; to: number} {
  const paras = blocks.filter((b): b is ParagraphBlock => b.type === 'p');
  let from = 0;
  for (let i = 0; i < Math.min(paras.length, TITLE_PAGE_MAX); i++) {
    const p = paras[i];
    if (i > 0 && p.pb) {
      from = i;
      break;
    }
    if (p.runs.some(r => r.pg)) {
      from = i + 1;
      break;
    }
  }
  const refs = findReferenceHeading(blocks);
  const to = refs ? paras.indexOf(refs) : paras.length;
  return {from: Math.min(from, to < 0 ? paras.length : to), to: to < 0 ? paras.length : to};
}

/** Words in the whole document and in its body text. */
export function wordCounts(blocks: Block[]): {all: number; body: number} {
  const paras = blocks.filter((b): b is ParagraphBlock => b.type === 'p');
  const {from, to} = bodyRange(blocks);
  let all = 0;
  let body = 0;
  paras.forEach((p, i) => {
    const n = paragraphWords(p);
    all += n;
    if (i >= from && i < to) {
      body += n;
    }
  });
  return {all, body};
}

/** "1,234 / 2,500 words" for the bottom line. */
export function targetLabel(body: number, target: number): string {
  return `${body.toLocaleString()} / ${target.toLocaleString()} words`;
}
