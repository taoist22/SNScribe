// Spelling: which words of a paragraph to check, and which to mark. The dictionary itself is
// native (SpellChecker.kt); this decides what counts as a word worth checking.

import {paragraphText, type ParagraphBlock} from '../model/docx';

export type Token = {start: number; end: number; word: string};

// Letters (Latin, with accents), with inner apostrophes: don't, children's, O’Brien.
const WORD = /[A-Za-zÀ-ɏ]+(?:['’][A-Za-zÀ-ɏ]+)*/g;
// Not words: web and e-mail addresses, DOIs, file names.
const SKIP = /\b(?:https?:\/\/|www\.)\S+|\b10\.\d{4,9}\/\S+|\S+@\S+\.\S+|\b\S+\.(?:pdf|docx?|epub|png|jpe?g|txt)\b/gi;

/** The words of `text` worth checking: not in addresses or DOIs, not abbreviations in capitals, not next to digits, not single letters. */
export function tokens(text: string): Token[] {
  const skipped: Array<[number, number]> = [];
  for (let m = SKIP.exec(text); m; m = SKIP.exec(text)) {
    skipped.push([m.index, m.index + m[0].length]);
  }
  SKIP.lastIndex = 0;
  const out: Token[] = [];
  for (let m = WORD.exec(text); m; m = WORD.exec(text)) {
    const start = m.index;
    const end = start + m[0].length;
    const word = m[0].replace(/['’]$/, '');
    if (word.length < 2 || skipped.some(([a, b]) => start < b && end > a)) {
      continue;
    }
    if (word.length >= 2 && word === word.toUpperCase()) {
      continue; // APA, RA, FIFO …
    }
    if (/[0-9_]/.test(text[start - 1] ?? '') || /[0-9_]/.test(text[end] ?? '')) {
      continue; // 2nd, H2O, B12 …
    }
    out.push({start, end: start + word.length, word});
  }
  WORD.lastIndex = 0;
  return out;
}

/** Normal form for the dictionary and the word lists: curly apostrophes straight. */
export const normal = (w: string) => w.replace(/’/g, "'");

/**
 * The ranges of `p` to mark: words the dictionary said are wrong (`bad`), unless added
 * (`mine`, lower-cased) or ignored; never the word being typed at `typingAt`.
 */
export function misspelledRanges(
  p: ParagraphBlock,
  bad: (word: string) => boolean,
  mine: Set<string>,
  ignored: Set<string>,
  typingAt?: number,
): Array<{start: number; end: number}> {
  return tokens(paragraphText(p))
    .filter(t => {
      const w = normal(t.word);
      return bad(w) && !mine.has(w.toLowerCase()) && !ignored.has(w) && t.end !== typingAt;
    })
    .map(({start, end}) => ({start, end}));
}
