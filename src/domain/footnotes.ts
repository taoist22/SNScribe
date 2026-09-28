// Editing a footnote's text in a plain text box, keeping its formatting where the words did
// not change: a Chicago note's italic book title stays italic when only the page is retyped.

import type {NotePiece} from '../model/docx';

export const noteText = (pieces: NotePiece[]) => pieces.map(p => p.t).join('');

/** The pieces covering [start, end) of the text. */
function slice(pieces: NotePiece[], start: number, end: number): NotePiece[] {
  const out: NotePiece[] = [];
  let at = 0;
  for (const p of pieces) {
    const s = Math.max(start, at);
    const e = Math.min(end, at + p.t.length);
    if (e > s) {
      out.push({...p, t: p.t.slice(s - at, e - at)});
    }
    at += p.t.length;
  }
  return out;
}

/** Joins neighbours that look the same. */
function tidy(pieces: NotePiece[]): NotePiece[] {
  const out: NotePiece[] = [];
  for (const p of pieces) {
    const last = out[out.length - 1];
    if (last && !!last.i === !!p.i && !!last.b === !!p.b) {
      last.t += p.t;
    } else if (p.t) {
      out.push({...p});
    }
  }
  return out;
}

/**
 * `old` with its text changed to `text`: the unchanged beginning and end keep their
 * formatting; what was typed in between takes the look of what it replaced (when that had
 * one look) or of the character before it.
 */
export function editPieces(old: NotePiece[], text: string): NotePiece[] {
  const before = noteText(old);
  let pre = 0;
  while (pre < before.length && pre < text.length && before[pre] === text[pre]) {
    pre++;
  }
  let suf = 0;
  while (suf < before.length - pre && suf < text.length - pre && before[before.length - 1 - suf] === text[text.length - 1 - suf]) {
    suf++;
  }
  const replaced = slice(old, pre, before.length - suf);
  const typed = text.slice(pre, text.length - suf);
  // Typed over text of one look: that look. Typed in without replacing: the look of the
  // character before (as typing in Word does), else of the one after.
  const around = slice(old, Math.max(0, pre - 1), pre)[0] ?? slice(old, pre, pre + 1)[0];
  const look: {i?: boolean; b?: boolean} =
    replaced.length > 0 ? (replaced.every(p => !!p.i === !!replaced[0].i && !!p.b === !!replaced[0].b) ? replaced[0] : {}) : around ?? {};
  return tidy([
    ...slice(old, 0, pre),
    ...(typed ? [{t: typed, ...(look.i ? {i: true} : {}), ...(look.b ? {b: true} : {})}] : []),
    ...slice(old, before.length - suf, before.length),
  ]);
}
