import {applyOps} from '../src/domain/edits';
import {htmlToPieces, inText, piecesText, referenceOps, type Source} from '../src/domain/citations';
import {paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

// Shapes as the Zotero Web API returned them (group 3470, 2026-09-27).
const BIB = `<div class="csl-bib-body" style="line-height: 2; padding-left: 1em; text-indent:-1em;">
  <div class="csl-entry">Smith, J., &amp; Lee, K. (2020). Mindfulness and test anxiety. <i>Journal of Educational Psychology</i>, <i>112</i>(3), 455&#x2013;467.</div>
</div>`;
const src: Source = {key: 'K1', title: 'Mindfulness', authors: 'Smith and Lee', year: '2020', citation: '<span>(Smith &amp; Lee, 2020)</span>', bibHtml: BIB};

const p = (index: number, t: string, kind: ParagraphBlock['kind'] = 'body'): ParagraphBlock => ({type: 'p', index, style: '', kind, level: kind === 'heading' ? 1 : 0, align: 'left', indent: 0, runs: t ? [{t}] : []});

describe('Zotero HTML', () => {
  it('keeps italics and decodes entities', () => {
    const pieces = htmlToPieces(BIB);
    expect(piecesText(pieces)).toBe('Smith, J., & Lee, K. (2020). Mindfulness and test anxiety. Journal of Educational Psychology, 112(3), 455–467.');
    expect(pieces.filter(x => x.i).map(x => x.t)).toEqual(['Journal of Educational Psychology', '112']);
  });
});

describe('in-text citations', () => {
  it('APA', () => {
    expect(inText(src, 'apa', false)).toBe('(Smith & Lee, 2020)');
    expect(inText(src, 'apa', false, '14')).toBe('(Smith & Lee, 2020, p. 14)');
    expect(inText(src, 'apa', false, '14-16')).toBe('(Smith & Lee, 2020, pp. 14–16)');
    expect(inText(src, 'apa', true)).toBe('Smith and Lee (2020)');
    expect(inText(src, 'apa', true, '14')).toBe('Smith and Lee (2020, p. 14)');
  });
  it('MLA and Chicago', () => {
    const mla = {...src, citation: '<span>(Smith and Lee)</span>'};
    expect(inText(mla, 'mla', false, '14')).toBe('(Smith and Lee 14)');
    expect(inText(mla, 'mla', true, '14')).toBe('Smith and Lee (14)');
    const chi = {...src, citation: '<span>(Smith and Lee 2020)</span>'};
    expect(inText(chi, 'chicago', false, '14')).toBe('(Smith and Lee 2020, 14)');
  });
});

describe('the reference list', () => {
  const entry = htmlToPieces(BIB);

  it('is created on a new page after the last paragraph', () => {
    const blocks: Block[] = [p(0, 'Body text.')];
    const {ops, added} = referenceOps(blocks, entry, 'apa');
    expect(added).toBe(true);
    const out = applyOps(blocks, ops) as ParagraphBlock[];
    expect(out.map(paragraphText)).toEqual(['Body text.', 'References', piecesText(entry)]);
    expect(out[1]).toMatchObject({align: 'center', pb: true});
    expect(out[2]).toMatchObject({first: -720, line: 480});
    expect(out[2].runs.filter(r => r.i).map(r => r.t)).toEqual(['Journal of Educational Psychology', '112']);
  });

  it('keeps entries in alphabetical order and never twice', () => {
    const blocks: Block[] = [p(0, 'Body.'), p(1, 'References'), p(2, 'Adams, A. (2019). A.'), p(3, 'Young, Y. (2018). Y.')];
    const out = applyOps(blocks, referenceOps(blocks, entry, 'apa').ops) as ParagraphBlock[];
    expect(out.map(paragraphText)).toEqual(['Body.', 'References', 'Adams, A. (2019). A.', piecesText(entry), 'Young, Y. (2018). Y.']);
    expect(referenceOps(out, entry, 'apa')).toEqual({ops: [], added: false});
  });

  it('goes first when it sorts before every entry', () => {
    const blocks: Block[] = [p(0, 'References'), p(1, 'Young, Y. (2018). Y.')];
    const out = applyOps(blocks, referenceOps(blocks, entry, 'apa').ops) as ParagraphBlock[];
    expect(out.map(paragraphText)).toEqual(['References', piecesText(entry), 'Young, Y. (2018). Y.']);
  });
});

describe('entries never inherit the heading', () => {
  it('no page break, bold or italic from the paragraph they are split from', () => {
    const heading: ParagraphBlock = {...p(0, ''), runs: [{t: 'References', b: true}], align: 'center', pb: true};
    const out = applyOps([heading], referenceOps([heading], htmlToPieces(BIB), 'apa').ops) as ParagraphBlock[];
    expect(out[1].pb).toBeFalsy();
    expect(out[1].align).toBe('left');
    expect(out[1].runs.filter(r => r.b).length).toBe(0);
  });
});

import {credentialsIn} from '../src/services/zotero';

describe('Zotero key from a text file', () => {
  it('finds the key and the user ID however the file is laid out', () => {
    expect(credentialsIn('P9NiFoyLeZu2bZNvvuQPDWsd')).toEqual({apiKey: 'P9NiFoyLeZu2bZNvvuQPDWsd'});
    expect(credentialsIn('Zotero\nuserID: 1234567\nkey: P9NiFoyLeZu2bZNvvuQPDWsd\n')).toEqual({apiKey: 'P9NiFoyLeZu2bZNvvuQPDWsd', userId: '1234567'});
    expect(credentialsIn('no key here 1234567')).toEqual({userId: '1234567'});
  });
});
