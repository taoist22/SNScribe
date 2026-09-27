import {htmlToPieces, inText, piecesText} from '../src/domain/citations';
import {cleanQuote, detailsFromCsl, detailsFromEpub, findDoi, formatSource, type Details} from '../src/domain/reference';

// doi.org's CSL JSON for 10.1037/0003-066X.59.1.29 (fetched 2026-09-27).
const csl = {
  type: 'journal-article',
  title: 'How the Mind Hurts and Heals the Body.',
  'container-title': 'American Psychologist',
  author: [{given: 'Oakley', family: 'Ray'}],
  issued: {'date-parts': [[2004]]},
  volume: '59',
  issue: '1',
  page: '29-40',
  publisher: 'American Psychological Association (APA)',
  DOI: '10.1037/0003-066x.59.1.29',
};

const text = (d: Details, style: 'apa' | 'mla' | 'chicago') => piecesText(htmlToPieces(formatSource(d, style).bibHtml));

describe('references for sources outside Zotero', () => {
  const article = detailsFromCsl(csl);

  it('reads doi.org details', () => {
    expect(article).toMatchObject({type: 'article', year: '2004', container: 'American Psychologist', pages: '29-40'});
  });

  it('APA 7 article: text, italics, in-text', () => {
    expect(text(article, 'apa')).toBe('Ray, O. (2004). How the Mind Hurts and Heals the Body. American Psychologist, 59(1), 29–40. https://doi.org/10.1037/0003-066x.59.1.29');
    const italic = htmlToPieces(formatSource(article, 'apa').bibHtml).filter(p => p.i).map(p => p.t);
    expect(italic).toEqual(['American Psychologist', '59']);
    const src = formatSource(article, 'apa');
    expect(inText(src, 'apa', false, '31')).toBe('(Ray, 2004, p. 31)');
    expect(inText(src, 'apa', true)).toBe('Ray (2004)');
  });

  it('two and three authors', () => {
    const two = {...article, authors: [{family: 'Smith', given: 'Jane'}, {family: 'Lee', given: 'Kim'}]};
    expect(inText(formatSource(two, 'apa'), 'apa', false)).toBe('(Smith & Lee, 2004)');
    expect(text(two, 'apa').startsWith('Smith, J., & Lee, K. (2004).')).toBe(true);
    expect(text(two, 'mla').startsWith('Smith, Jane, and Kim Lee. “How the Mind Hurts and Heals the Body.” American Psychologist, vol. 59, no. 1, 2004, pp. 29–40.')).toBe(true);
    const three = {...two, authors: [...two.authors, {family: 'Wu', given: 'An'}]};
    expect(inText(formatSource(three, 'apa'), 'apa', false)).toBe('(Smith et al., 2004)');
  });

  it('Chicago author-date', () => {
    expect(text(article, 'chicago')).toBe('Ray, Oakley. 2004. “How the Mind Hurts and Heals the Body.” American Psychologist 59 (1): 29–40. https://doi.org/10.1037/0003-066x.59.1.29.');
    expect(inText(formatSource(article, 'chicago'), 'chicago', false, '31')).toBe('(Ray 2004, 31)');
  });

  it('a book from an EPUB', () => {
    const book = detailsFromEpub({title: 'Financial Accounting', creators: ['Jane Q. Smith'], date: '2021-05-01', publisher: 'OpenStax'});
    expect(text(book, 'apa')).toBe('Smith, J. Q. (2021). Financial Accounting. OpenStax.');
    expect(htmlToPieces(formatSource(book, 'apa').bibHtml).find(p => p.i)?.t).toBe('Financial Accounting');
    expect(text(book, 'mla')).toBe('Smith, Jane Q. Financial Accounting. OpenStax, 2021.');
  });
});

describe('quotes from PDFs', () => {
  it('finds a DOI on a first page', () => {
    expect(findDoi('Copyright … https://doi.org/10.1037/0003-066X.59.1.29. Received')).toBe('10.1037/0003-066X.59.1.29');
    expect(findDoi('no identifier')).toBeUndefined();
  });

  it('joins line-break hyphens and collapses lines', () => {
    expect(cleanQuote('increases inflam-\nmation and the\n risk of RA ')).toBe('increases inflammation and the risk of RA');
    expect(cleanQuote('a well-known effect')).toBe('a well-known effect');
  });
});
