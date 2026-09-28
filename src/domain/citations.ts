// Citations from a Zotero library: in-text citations and reference-list entries, formatted
// by Zotero itself (its Web API returns them as HTML in any CSL style), turned into the
// plugin's ordinary edits: text, paragraph splits, italics and paragraph settings.

import {paragraphText, type Block, type ParagraphBlock} from '../model/docx';
import type {Op} from './edits';

export type CiteStyle = 'apa' | 'mla' | 'chicago' | 'chicago-notes';

export const CITE_STYLES: Array<{id: CiteStyle; name: string; csl: string; heading: string}> = [
  {id: 'apa', name: 'APA 7', csl: 'apa', heading: 'References'},
  {id: 'mla', name: 'MLA 9', csl: 'modern-language-association', heading: 'Works Cited'},
  {id: 'chicago', name: 'Chicago (author-date)', csl: 'chicago-author-date', heading: 'References'},
  {id: 'chicago-notes', name: 'Chicago (notes)', csl: 'chicago-note-bibliography', heading: 'Bibliography'},
];

/** Styles that cite in footnotes rather than in the sentence. */
export const isNoteStyle = (style: CiteStyle) => style === 'chicago-notes';

/** One source from the library, with Zotero's formatting for the chosen style. */
export type Source = {
  key: string;
  title: string;
  /** "Smith", "Smith and Lee", "Smith et al." (Zotero's creator summary). */
  authors: string;
  year: string;
  /** Zotero's in-text citation, e.g. "(Smith & Lee, 2020)". */
  citation: string;
  /** Zotero's reference entry as HTML (italics as <i>). */
  bibHtml: string;
  /** A book-length work (title in italics) or a shorter one (title in quotation marks): for short notes. */
  kind?: 'book' | 'article';
};

/** A piece of formatted text. */
export type Piece = {t: string; i?: boolean; b?: boolean};

const ENTITIES: Record<string, string> = {amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' '};

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') {
      const n = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(n) ? String.fromCodePoint(n) : m;
    }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}

/**
 * Zotero's HTML (a csl-entry, or a citation span) as text pieces: <i>/<em> italic,
 * <b>/<strong> bold, other tags dropped, entities decoded, whitespace runs collapsed.
 */
export function htmlToPieces(html: string): Piece[] {
  const out: Piece[] = [];
  let italic = 0;
  let bold = 0;
  const push = (t: string) => {
    if (!t) {
      return;
    }
    const last = out[out.length - 1];
    const p: Piece = {t, ...(italic ? {i: true} : {}), ...(bold ? {b: true} : {})};
    if (last && !!last.i === !!p.i && !!last.b === !!p.b) {
      last.t += t;
    } else {
      out.push(p);
    }
  };
  const re = /<\/?([a-z0-9]+)[^>]*>|([^<]+)/gi;
  for (let m = re.exec(html); m; m = re.exec(html)) {
    if (m[2] !== undefined) {
      push(decodeEntities(m[2]).replace(/\s+/g, ' '));
      continue;
    }
    const tag = m[1].toLowerCase();
    const closing = m[0][1] === '/';
    if (tag === 'i' || tag === 'em') {
      italic += closing ? -1 : 1;
    } else if (tag === 'b' || tag === 'strong') {
      bold += closing ? -1 : 1;
    }
  }
  // Trim the ends (the HTML is indented).
  if (out.length) {
    out[0].t = out[0].t.replace(/^\s+/, '');
    out[out.length - 1].t = out[out.length - 1].t.replace(/\s+$/, '');
  }
  return out.filter(p => p.t.length > 0);
}

export const piecesText = (pieces: Piece[]) => pieces.map(p => p.t).join('');

/** A title shortened for a short note (Chicago): up to its colon, at most four words. */
export function shortTitle(title: string): string {
  const main = title.split(/[:?]/)[0].trim();
  const words = main.split(/\s+/).filter(Boolean);
  return words.length > 4 ? words.slice(0, 4).join(' ') : main;
}

/** Joins neighbouring pieces that look the same. */
function tidyPieces(pieces: Piece[]): Piece[] {
  const out: Piece[] = [];
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
 * A footnote citing `source` (Chicago notes): the full note the first time, a short note
 * ("Smith, Short Title, 23.") after that. The page goes where Chicago puts it: after the
 * publication details of a book, in place of an article's page range.
 */
export function notePieces(source: Source, page: string, short: boolean): Piece[] {
  const pg = page.trim().replace(/\s*-\s*/g, '–');
  if (short) {
    const t = shortTitle(source.title);
    const who = source.authors ? `${source.authors}, ` : '';
    if (source.kind === 'article') {
      return tidyPieces([{t: `${who}“${t}${pg ? `,” ${pg}.` : '.”'}`}]);
    }
    return tidyPieces([{t: who}, {t, i: true}, {t: pg ? `, ${pg}.` : '.'}]);
  }
  const pieces = htmlToPieces(source.citation).map(p => ({...p}));
  if (pieces.length === 0) {
    return [];
  }
  const last = pieces[pieces.length - 1];
  last.t = last.t.replace(/\s+$/, '');
  if (pg) {
    if (source.kind === 'article' && /: ?[0-9ivxlc]+(?:[–-][0-9ivxlc]+)?\.$/i.test(last.t)) {
      last.t = last.t.replace(/: ?[0-9ivxlc]+(?:[–-][0-9ivxlc]+)?\.$/i, `: ${pg}.`);
    } else if (last.t.endsWith('.')) {
      last.t = `${last.t.slice(0, -1)}, ${pg}.`;
    } else {
      last.t = `${last.t}, ${pg}.`;
    }
  } else if (!/[.!?]$/.test(last.t)) {
    last.t += '.';
  }
  return tidyPieces(pieces);
}

/** "14" → "p. 14"; "14-16" / "14–16" → "pp. 14–16" (APA). */
function apaPages(page: string): string {
  const p = page.trim().replace(/\s*-\s*/g, '–');
  return /[–,]/.test(p) ? `pp. ${p}` : `p. ${p}`;
}

/**
 * The in-text citation, parenthetical or narrative, with an optional page:
 * APA (Smith & Lee, 2020, p. 14) / Smith and Lee (2020, p. 14);
 * MLA (Smith and Lee 14) / Smith and Lee (14);
 * Chicago author-date (Smith and Lee 2020, 14) / Smith and Lee (2020, 14).
 */
export function inText(source: Source, style: CiteStyle, narrative: boolean, page = ''): string {
  if (isNoteStyle(style)) {
    // Shown for checking: the footnote's text.
    return piecesText(notePieces(source, page, false));
  }
  const pg = page.trim();
  const zotero = piecesText(htmlToPieces(source.citation));
  if (!narrative) {
    if (!pg) {
      return zotero;
    }
    const at = zotero.lastIndexOf(')');
    if (at < 0) {
      return zotero;
    }
    const add = style === 'apa' ? `, ${apaPages(pg)}` : style === 'mla' ? ` ${pg}` : `, ${pg}`;
    return zotero.slice(0, at) + add + zotero.slice(at);
  }
  // Narrative: the authors in the sentence, the rest in parentheses. With no authors,
  // Zotero's citation without its parentheses.
  if (!source.authors) {
    return zotero.replace(/^\(|\)$/g, '');
  }
  const inside =
    style === 'apa' ? [source.year, pg ? apaPages(pg) : ''] : style === 'mla' ? [pg] : [source.year, pg];
  const paren = inside.filter(Boolean).join(', ');
  return paren ? `${source.authors} (${paren})` : source.authors;
}

/** The reference-list heading paragraph ("References", "Works Cited", "Bibliography"), or null. */
export function findReferenceHeading(blocks: Block[]): ParagraphBlock | null {
  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i];
    if (b.type === 'p' && /^(references|works cited|bibliography|reference list)$/i.test(paragraphText(b).trim())) {
      return b;
    }
  }
  return null;
}

/** The entries under the reference heading: body paragraphs up to the next heading. */
export function referenceEntries(blocks: Block[], heading: ParagraphBlock): ParagraphBlock[] {
  const out: ParagraphBlock[] = [];
  const start = blocks.indexOf(heading);
  for (let i = start + 1; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.type !== 'p') {
      continue;
    }
    if (b.kind !== 'body') {
      break;
    }
    out.push(b);
  }
  return out;
}

const DOUBLE = 480;
const HANGING = -720;

/** Italics (and bold) of a formatted entry written into paragraph `para` from `offset`. */
function formatOps(para: number, offset: number, pieces: Piece[]): Op[] {
  const ops: Op[] = [];
  let at = offset;
  for (const p of pieces) {
    if (p.i) {
      ops.push({op: 'format', para, start: at, end: at + p.t.length, prop: 'i', on: true});
    }
    if (p.b) {
      ops.push({op: 'format', para, start: at, end: at + p.t.length, prop: 'b', on: true});
    }
    at += p.t.length;
  }
  return ops;
}

/**
 * The edits that put a source's entry in the reference list: alphabetically among the
 * entries under the heading, creating the heading (on a new page, centred, bold) after the
 * last paragraph when there is none. Nothing when the entry is already there.
 */
export function referenceOps(blocks: Block[], entry: Piece[], style: CiteStyle): {ops: Op[]; added: boolean} {
  const text = piecesText(entry);
  const paras = blocks.filter((b): b is ParagraphBlock => b.type === 'p');
  const last = paras[paras.length - 1];
  if (!last) {
    return {ops: [], added: false};
  }
  // Split off the heading or another entry, a new paragraph brings their settings along:
  // each is set explicitly (no page break, no bold or italic but the entry's own).
  const entryPara = (para: number): Op[] => [
    {op: 'text', para, start: 0, end: 0, text},
    {op: 'style', para, kind: 'normal'},
    {op: 'para', para, align: 'left', line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: HANGING, pb: false},
    {op: 'format', para, start: 0, end: text.length, prop: 'b', on: false},
    {op: 'format', para, start: 0, end: text.length, prop: 'i', on: false},
    ...formatOps(para, 0, entry),
  ];
  const heading = findReferenceHeading(blocks);
  if (!heading) {
    const lastLen = paragraphText(last).length;
    const h = last.index + 1;
    const title = CITE_STYLES.find(s => s.id === style)!.heading;
    return {
      ops: [
        {op: 'split', para: last.index, offset: lastLen},
        {op: 'text', para: h, start: 0, end: 0, text: title},
        {op: 'style', para: h, kind: 'normal'},
        {op: 'para', para: h, align: 'center', line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0, pb: true},
        {op: 'format', para: h, start: 0, end: title.length, prop: 'b', on: style !== 'mla'},
        {op: 'format', para: h, start: 0, end: title.length, prop: 'i', on: false},
        {op: 'split', para: h, offset: title.length},
        ...entryPara(h + 1),
      ],
      added: true,
    };
  }
  const entries = referenceEntries(blocks, heading);
  if (entries.some(e => paragraphText(e).trim() === text.trim())) {
    return {ops: [], added: false};
  }
  // An empty entry paragraph (the list's placeholder) is used as it is.
  const empty = entries.find(e => paragraphText(e).trim() === '');
  if (empty) {
    return {ops: entryPara(empty.index), added: true};
  }
  const key = text.toLocaleLowerCase();
  const before = [...entries].reverse().find(e => paragraphText(e).toLocaleLowerCase() <= key);
  const after = before ?? heading;
  const len = paragraphText(after).length;
  const p = after.index + 1;
  return {ops: [{op: 'split', para: after.index, offset: len}, ...entryPara(p)], added: true};
}
