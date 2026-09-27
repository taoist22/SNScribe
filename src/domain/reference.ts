// References formatted by the plugin itself, for sources not in Zotero: details from a DOI
// (doi.org's CSL JSON), from an EPUB's own metadata, or typed in. Covers the common cases —
// journal article, book, anything else (web page, report) — in APA 7, MLA 9 and Chicago
// author-date. The result has the same shape as a Zotero-formatted source, so citing and the
// reference list treat both alike.

import type {CiteStyle, Source} from './citations';

export type Person = {family: string; given?: string};

export type Details = {
  type: 'article' | 'book' | 'other';
  authors: Person[];
  year?: string;
  title: string;
  /** Journal (article) or website / larger work (other). */
  container?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  publisher?: string;
  doi?: string;
  url?: string;
};

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const i = (s: string) => `<i>${esc(s)}</i>`;
const clean = (s?: string) => (s ?? '').trim().replace(/\s+/g, ' ');
/** A title without its own closing full stop (the formats add their own). */
const title = (s: string) => clean(s).replace(/[.]+$/, '');
const dash = (pages: string) => clean(pages).replace(/\s*-+\s*/g, '–');
const doiUrl = (doi: string) => `https://doi.org/${clean(doi).replace(/^https?:\/\/(dx\.)?doi\.org\//i, '').replace(/^doi:\s*/i, '')}`;

/** "Oakley" → "O.", "Mary Ann" → "M. A.", "Jean-Paul" → "J.-P." */
function initials(given?: string): string {
  return clean(given)
    .split(/\s+/)
    .filter(Boolean)
    .map(part =>
      part
        .split('-')
        .map(p => (p.endsWith('.') && p.length <= 3 ? p : `${p[0].toUpperCase()}.`))
        .join('-'),
    )
    .join(' ');
}

/** A full stop after `s`, unless it already ends with one (a name ending in an initial). */
const stop = (s: string) => (s.endsWith('.') ? s : `${s}.`);
const full = (p: Person) => (p.given ? `${clean(p.given)} ${clean(p.family)}` : clean(p.family));
const inverted = (p: Person) => (p.given ? `${clean(p.family)}, ${clean(p.given)}` : clean(p.family));

function apaNames(authors: Person[]): string {
  const names = authors.map(p => (p.given ? `${clean(p.family)}, ${initials(p.given)}` : clean(p.family)));
  if (names.length <= 1) {
    return names[0] ?? '';
  }
  if (names.length > 20) {
    return `${names.slice(0, 19).join(', ')}, . . . ${names[names.length - 1]}`;
  }
  return `${names.slice(0, -1).join(', ')}, & ${names[names.length - 1]}`;
}

function mlaNames(authors: Person[]): string {
  if (authors.length === 0) {
    return '';
  }
  if (authors.length === 1) {
    return inverted(authors[0]);
  }
  if (authors.length === 2) {
    return `${inverted(authors[0])}, and ${full(authors[1])}`;
  }
  return `${inverted(authors[0])}, et al.`;
}

function chicagoNames(authors: Person[]): string {
  if (authors.length === 0) {
    return '';
  }
  if (authors.length === 1) {
    return inverted(authors[0]);
  }
  const rest = authors.slice(1).map(full);
  return `${inverted(authors[0])}, ${rest.slice(0, -1).map(n => `${n}, `).join('')}and ${rest[rest.length - 1]}`;
}

/** Who, as the sentence names them: "Ray", "Smith and Lee", "Smith et al." */
function summary(authors: Person[], style: CiteStyle): string {
  const fam = authors.map(p => clean(p.family));
  if (fam.length === 0) {
    return '';
  }
  if (fam.length === 1) {
    return fam[0];
  }
  if (fam.length === 2) {
    return `${fam[0]} and ${fam[1]}`;
  }
  if (style === 'chicago' && fam.length === 3) {
    return `${fam[0]}, ${fam[1]}, and ${fam[2]}`;
  }
  return `${fam[0]} et al.`;
}

/** A source formatted in `style`, in the shape Zotero's formatting has. */
export function formatSource(d: Details, style: CiteStyle, key = ''): Source {
  const year = clean(d.year) || (style === 'apa' ? 'n.d.' : '');
  const who = summary(d.authors, style);
  const t = title(d.title);
  const container = clean(d.container);
  const link = d.doi ? doiUrl(d.doi) : clean(d.url);
  let bib = '';
  let cite = '';
  if (style === 'apa') {
    const names = apaNames(d.authors);
    const lead = names ? `${esc(names)} (${esc(year)}). ` : '';
    const head = names ? '' : `${d.type === 'article' ? esc(t) : i(t)}. (${esc(year)}). `;
    if (d.type === 'article') {
      const vol = d.volume ? `, ${i(clean(d.volume))}${d.issue ? `(${esc(clean(d.issue))})` : ''}` : '';
      const pages = d.pages ? `, ${esc(dash(d.pages))}` : '';
      bib = `${lead || head}${names ? `${esc(t)}. ` : ''}${container ? i(container) : ''}${vol}${pages}.`;
    } else if (d.type === 'book') {
      bib = `${lead || head}${names ? `${i(t)}. ` : ''}${d.publisher ? `${esc(clean(d.publisher))}.` : ''}`;
    } else {
      bib = `${lead || head}${names ? `${i(t)}. ` : ''}${container ? `${esc(container)}.` : ''}`;
    }
    bib = `${bib.trim()}${link ? ` ${esc(link)}` : ''}`;
    cite = `(${esc(who || `“${t}”`)}, ${esc(year)})`.replace(' and ', ' &amp; ');
  } else if (style === 'mla') {
    const names = mlaNames(d.authors);
    const lead = names ? `${esc(stop(names))} ` : '';
    if (d.type === 'article') {
      const parts = [container ? i(container) : '', d.volume ? `vol. ${esc(clean(d.volume))}` : '', d.issue ? `no. ${esc(clean(d.issue))}` : '', esc(year), d.pages ? `pp. ${esc(dash(d.pages))}` : ''].filter(Boolean);
      bib = `${lead}“${esc(t)}.” ${parts.join(', ')}.`;
    } else if (d.type === 'book') {
      bib = `${lead}${i(t)}. ${[d.publisher ? esc(clean(d.publisher)) : '', esc(year)].filter(Boolean).join(', ')}.`;
    } else {
      bib = `${lead}${i(t)}. ${[container ? esc(container) : '', esc(year)].filter(Boolean).join(', ')}.`;
    }
    bib = `${bib.replace(/\s+\./, '.').trim()}${link ? ` ${esc(link)}.` : ''}`;
    cite = `(${esc(who || `“${t}”`)})`;
  } else {
    const names = chicagoNames(d.authors);
    const lead = names ? `${esc(stop(names))} ${esc(year)}. ` : '';
    if (d.type === 'article') {
      const vol = d.volume ? ` ${esc(clean(d.volume))}${d.issue ? ` (${esc(clean(d.issue))})` : ''}` : '';
      const pages = d.pages ? `: ${esc(dash(d.pages))}` : '';
      bib = `${lead}“${esc(t)}.” ${container ? i(container) : ''}${vol}${pages}.`;
    } else if (d.type === 'book') {
      bib = `${lead}${i(t)}. ${d.publisher ? `${esc(clean(d.publisher))}.` : ''}`;
    } else {
      bib = `${lead}${i(t)}. ${container ? `${esc(container)}.` : ''}`;
    }
    bib = `${bib.trim()}${link ? ` ${esc(link)}.` : ''}`;
    cite = `(${esc(who || `“${t}”`)} ${esc(year)})`;
  }
  return {key, title: t, authors: who, year, citation: `<span>${cite}</span>`, bibHtml: `<div class="csl-entry">${bib}</div>`};
}

/** doi.org's CSL JSON (Accept: application/vnd.citationstyles.csl+json) as details. */
export function detailsFromCsl(csl: Record<string, unknown>): Details {
  const str = (v: unknown) => (Array.isArray(v) ? String(v[0] ?? '') : v === undefined || v === null ? '' : String(v));
  const people = (v: unknown): Person[] =>
    Array.isArray(v)
      ? v
          .map(p => p as {family?: string; given?: string; literal?: string})
          .map(p => ({family: p.family ?? p.literal ?? '', ...(p.given ? {given: p.given} : {})}))
          .filter(p => p.family)
      : [];
  const type = str(csl.type);
  const issued = (csl.issued as {'date-parts'?: number[][]} | undefined)?.['date-parts']?.[0]?.[0];
  const authors = people(csl.author);
  return {
    type: type === 'journal-article' || type === 'article-journal' ? 'article' : type === 'book' || type === 'monograph' ? 'book' : 'other',
    authors: authors.length ? authors : people(csl.editor),
    year: issued ? String(issued) : undefined,
    title: str(csl.title),
    container: str(csl['container-title']) || undefined,
    volume: str(csl.volume) || undefined,
    issue: str(csl.issue) || undefined,
    pages: str(csl.page) || undefined,
    publisher: str(csl.publisher) || undefined,
    doi: str(csl.DOI) || undefined,
  };
}

/** An EPUB's own metadata (DocxModule.epubInfo) as details: a book. */
export function detailsFromEpub(info: {title?: string; creators?: string[]; date?: string; publisher?: string}): Details {
  return {
    type: 'book',
    authors: (info.creators ?? []).map(c => {
      const s = clean(c);
      if (s.includes(',')) {
        const [family, given] = s.split(',', 2);
        return {family: clean(family), given: clean(given)};
      }
      const parts = s.split(' ');
      return parts.length > 1 ? {family: parts[parts.length - 1], given: parts.slice(0, -1).join(' ')} : {family: s};
    }),
    year: /\d{4}/.exec(info.date ?? '')?.[0],
    title: clean(info.title),
    publisher: clean(info.publisher) || undefined,
  };
}

/** A DOI in text (the first page of an article usually prints one). */
export function findDoi(text: string): string | undefined {
  const m = /\b(10\.\d{4,9}\/[^\s"<>]+)/i.exec(text);
  return m ? m[1].replace(/[.,;)\]]+$/, '') : undefined;
}

/** Selected PDF text made readable: line-break hyphens joined, line breaks and runs of spaces collapsed. */
export function cleanQuote(text: string): string {
  return text
    // Plain character ranges: the device's JavaScript engine may lack \p{…} classes.
    .replace(/([A-Za-z\u00C0-\u024F])-\s*\n\s*([a-z\u00DF-\u00FF])/g, '$1$2')
    .replace(/\s+/g, ' ')
    .trim();
}
