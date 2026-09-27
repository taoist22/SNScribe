// Handwritten notes: private to the Supernote, kept beside the document (never in the Word
// file). Each note is a picture of the ink plus an anchor — the words it was written about
// — so it can find its place again after edits, here or in Word elsewhere.

import {paragraphText, type Block, type ParagraphBlock} from '../model/docx';

export type InkNote = {
  id: string;
  /** The picture, in private storage. */
  path: string;
  width: number;
  height: number;
  /** Where it was written: paragraph index and text offset… */
  para: number;
  at: number;
  /** …and the words there (the selection, or the word at the caret); '' for an empty line. */
  quote: string;
  created: string;
};

export type Placed = {para: number; at: number};

/**
 * Where a note belongs in `blocks` now: its words at its own place, else the occurrence
 * of its words nearest its old paragraph, else (no words) its paragraph if it still
 * exists. Null: its words are gone — the note is listed as unplaced, never dropped.
 */
export function placeNote(note: InkNote, blocks: Block[]): Placed | null {
  const paras = blocks.filter((b): b is ParagraphBlock => b.type === 'p');
  const own = paras.find(p => p.index === note.para);
  if (!note.quote) {
    return own ? {para: own.index, at: Math.min(note.at, paragraphText(own).length)} : null;
  }
  if (own && paragraphText(own).slice(note.at, note.at + note.quote.length) === note.quote) {
    return {para: own.index, at: note.at};
  }
  let best: Placed | null = null;
  let bestScore = Infinity;
  for (const p of paras) {
    const text = paragraphText(p);
    for (let i = text.indexOf(note.quote); i >= 0; i = text.indexOf(note.quote, i + 1)) {
      // Nearest paragraph first, then nearest offset within it.
      const score = Math.abs(p.index - note.para) * 100000 + Math.abs(i - note.at);
      if (score < bestScore) {
        bestScore = score;
        best = {para: p.index, at: i};
      }
    }
  }
  return best;
}

/** The notes with their anchors moved to where they are now found (unplaced ones unchanged). */
export function reanchor(notes: InkNote[], blocks: Block[]): InkNote[] {
  return notes.map(n => {
    const at = placeNote(n, blocks);
    return at && (at.para !== n.para || at.at !== n.at) ? {...n, ...at} : n;
  });
}

export function parseNotes(json: string | null | undefined): InkNote[] {
  try {
    const list = JSON.parse(json ?? '[]');
    return Array.isArray(list) ? list.filter(n => n && typeof n.path === 'string' && typeof n.para === 'number') : [];
  } catch {
    return [];
  }
}
