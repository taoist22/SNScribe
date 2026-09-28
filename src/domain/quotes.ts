// Quotes kept from PDFs and EPUBs (the "Quote → SNScribe" button on their selection toolbar), and
// where each source file's citation details come from (chosen once per file, remembered).

import {paragraphText, type Block} from '../model/docx';
import {htmlToPieces, inText, isNoteStyle, notePieces, referenceOps, type CiteStyle, type Source} from './citations';
import type {Details} from './reference';
import type {Op} from './edits';

export type SavedQuote = {
  id: string;
  text: string;
  /** The PDF/EPUB it came from. */
  path: string;
  /** Its page (1 = first); EPUB pages depend on the reader's layout. */
  page: number;
  /** A DOI found on the document's first page, if any. */
  doi?: string;
  captured: string;
};

/** Where a source file's citation comes from. */
export type SourceRef = {kind: 'zotero'; key: string} | {kind: 'details'; details: Details};

export const QUOTES_KEY = 'quotes';
export const SOURCES_KEY = 'sources';

export function parseList<T>(json: string | null | undefined): T[] {
  try {
    const v = JSON.parse(json ?? '[]');
    return Array.isArray(v) ? (v as T[]) : [];
  } catch {
    return [];
  }
}

export function parseMap<T>(json: string | null | undefined): Record<string, T> {
  try {
    const v = JSON.parse(json ?? '{}');
    return v && typeof v === 'object' && !Array.isArray(v) ? (v as Record<string, T>) : {};
  } catch {
    return {};
  }
}

export const fileName = (path: string) => path.split('/').pop() ?? path;
export const isEpub = (path: string) => /\.epub$/i.test(path);

/** APA's line: 40 words or more is a block quotation (MLA and Chicago use similar lengths). */
export const BLOCK_WORDS = 40;
const words = (s: string) => s.trim().split(/\s+/).filter(Boolean).length;

/**
 * The edits that quote `text` at `at` with its citation, and add the source to the reference
 * list — one undo step. Short: in the sentence, in quotation marks, cited after the closing
 * mark. Long (BLOCK_WORDS+): a block quotation (Quote style) after the caret's paragraph, no
 * quotation marks, cited after its final punctuation.
 */
export function quoteOps(
  blocks: Block[],
  at: {para: number; offset: number},
  text: string,
  source: Source,
  style: CiteStyle,
  page: string,
  note?: {id: number; short: boolean},
): {ops: Op[]; block: boolean; caret: {para: number; offset: number}} {
  const quote = text.trim().replace(/^[“"]+|[”"]+$/g, '');
  // Chicago notes: no citation in the sentence; a footnote after the quotation instead.
  const footnote = isNoteStyle(style) && note ? {id: note.id, pieces: notePieces(source, page, note.short)} : null;
  const cite = footnote ? '' : inText(source, style, false, page);
  const ops: Op[] = [];
  let caret = at;
  const block = words(quote) >= BLOCK_WORDS;
  const p = blocks.find(b => b.type === 'p' && b.index === at.para);
  const ptext = p && p.type === 'p' ? paragraphText(p) : '';
  if (block) {
    const next = at.para + 1;
    const ended = /[.!?]$/.test(quote) ? quote : `${quote}.`;
    const body = cite ? `${ended} ${cite}` : ended;
    ops.push(
      {op: 'split', para: at.para, offset: ptext.length},
      {op: 'text', para: next, start: 0, end: 0, text: body},
      {op: 'style', para: next, kind: 'quote'},
      {op: 'para', para: next, first: 0, pb: false},
    );
    if (footnote) {
      ops.push({op: 'footnote', para: next, at: body.length, id: footnote.id, pieces: footnote.pieces});
    }
    caret = {para: next, offset: body.length + (footnote ? 1 : 0)};
  } else {
    let insert = cite ? `“${quote}” ${cite}` : `“${quote}”`;
    if (at.offset > 0 && !/[\s(\[]/.test(ptext[at.offset - 1])) {
      insert = ` ${insert}`;
    }
    if (at.offset < ptext.length && /[A-Za-z0-9À-ɏ“"]/.test(ptext[at.offset])) {
      insert = `${insert} `;
    }
    ops.push({op: 'text', para: at.para, start: at.offset, end: at.offset, text: insert});
    if (footnote) {
      // The number right after the closing quotation mark.
      const after = at.offset + insert.trimEnd().length;
      ops.push({op: 'footnote', para: at.para, at: after, id: footnote.id, pieces: footnote.pieces});
      caret = {para: at.para, offset: at.offset + insert.length + 1};
    } else {
      caret = {para: at.para, offset: at.offset + insert.length};
    }
  }
  return {ops, block, caret};
}

/** quoteOps plus the reference-list entry, computed on the document after the quote. */
export function quoteWithReference(
  blocks: Block[],
  applyOps: (b: Block[], o: Op[]) => Block[],
  at: {para: number; offset: number},
  text: string,
  source: Source,
  style: CiteStyle,
  page: string,
  note?: {id: number; short: boolean},
): {ops: Op[]; block: boolean; caret: {para: number; offset: number}; added: boolean} {
  const q = quoteOps(blocks, at, text, source, style, page, note);
  const ref = referenceOps(applyOps(blocks, q.ops), htmlToPieces(source.bibHtml), style);
  return {...q, ops: [...q.ops, ...ref.ops], added: ref.added};
}
