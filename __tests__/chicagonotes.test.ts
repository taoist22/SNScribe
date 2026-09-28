import {notePieces, piecesText, shortTitle, type Source} from '../src/domain/citations';
import {applyOps, expectedTexts} from '../src/domain/edits';
import {quoteWithReference} from '../src/domain/quotes';
import {formatSource} from '../src/domain/reference';
import {OBJECT, paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const book: Source = {
  key: 'K1',
  title: 'The Craft of Research: Fourth Edition',
  authors: 'Booth et al.',
  year: '2016',
  citation: '<span>Wayne C. Booth et al., <i>The Craft of Research: Fourth Edition</i> (University of Chicago Press, 2016).</span>',
  bibHtml: '<div class="csl-entry">Booth, Wayne C., et al. <i>The Craft of Research</i>. University of Chicago Press, 2016.</div>',
  kind: 'book',
};
const article: Source = {
  key: 'K2',
  title: 'Reading Closely',
  authors: 'Lee',
  year: '2020',
  citation: '<span>Ann Lee, “Reading Closely,” <i>Journal of Reading</i> 12, no. 3 (2020): 40–60.</span>',
  bibHtml: '<div class="csl-entry">Lee, Ann. “Reading Closely.” <i>Journal of Reading</i> 12, no. 3 (2020): 40–60.</div>',
  kind: 'article',
};

describe('Chicago notes', () => {
  it('full notes put the page where Chicago does', () => {
    expect(piecesText(notePieces(book, '23', false))).toBe('Wayne C. Booth et al., The Craft of Research: Fourth Edition (University of Chicago Press, 2016), 23.');
    expect(notePieces(book, '23', false).find(p => p.i)?.t).toBe('The Craft of Research: Fourth Edition');
    expect(piecesText(notePieces(article, '45', false))).toBe('Ann Lee, “Reading Closely,” Journal of Reading 12, no. 3 (2020): 45.');
    expect(piecesText(notePieces(article, '', false))).toBe('Ann Lee, “Reading Closely,” Journal of Reading 12, no. 3 (2020): 40–60.');
  });

  it('short notes after the first', () => {
    expect(shortTitle('The Craft of Research: Fourth Edition')).toBe('The Craft of Research');
    expect(piecesText(notePieces(book, '30', true))).toBe('Booth et al., The Craft of Research, 30.');
    expect(notePieces(book, '30', true)[1]).toEqual({t: 'The Craft of Research', i: true});
    expect(piecesText(notePieces(article, '41', true))).toBe('Lee, “Reading Closely,” 41.');
    expect(piecesText(notePieces(article, '', true))).toBe('Lee, “Reading Closely.”');
  });

  it('formats sources that are not in Zotero', () => {
    const s = formatSource({type: 'book', authors: [{family: 'Smith', given: 'Jane'}], year: '2020', title: 'A Book', publisher: 'Press'}, 'chicago-notes');
    expect(piecesText(notePieces(s, '5', false))).toBe('Jane Smith, A Book (Press, 2020), 5.');
    expect(s.bibHtml).toBe('<div class="csl-entry">Smith, Jane. <i>A Book</i>. Press, 2020.</div>');
    const a = formatSource({type: 'article', authors: [{family: 'Lee', given: 'Ann'}], year: '2020', title: 'Reading', container: 'Journal', volume: '12', issue: '3', pages: '40-60'}, 'chicago-notes');
    expect(piecesText(notePieces(a, '45', false))).toBe('Ann Lee, “Reading,” Journal 12, no. 3 (2020): 45.');
  });

  it('a quote gets a footnote after it, and the source goes in the bibliography', () => {
    const para = (index: number, t: string): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t}]});
    const blocks: Block[] = [para(0, 'As Booth says, and so on.')];
    const r = quoteWithReference(blocks, applyOps, {para: 0, offset: 14}, 'research is a craft', book, 'chicago-notes', '23', {id: 1, short: false});
    const after = applyOps(blocks, r.ops);
    expect(paragraphText(after[0] as ParagraphBlock)).toBe(`As Booth says, “research is a craft”${OBJECT} and so on.`);
    expect(texts(after).slice(-2)).toEqual(['Bibliography', 'Booth, Wayne C., et al. The Craft of Research. University of Chicago Press, 2016.']);
    expect(expectedTexts(blocks, r.ops)).toEqual(texts(after));
  });
});

const texts = (blocks: Block[]) => blocks.map(b => (b.type === 'p' ? paragraphText(b) : ''));
