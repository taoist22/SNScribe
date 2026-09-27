import {applyOps, headerFooterAfter, linkProblem, linkSpan, linkUrl} from '../src/domain/edits';
import {presetOps} from '../src/domain/presets';
import {OBJECT, paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const para = (index: number, runs: ParagraphBlock['runs']): ParagraphBlock => ({
  type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs,
});

describe('linkUrl', () => {
  it('turns DOIs into doi.org addresses', () => {
    expect(linkUrl('10.1037/0000165-000')).toBe('https://doi.org/10.1037/0000165-000');
    expect(linkUrl('doi:10.1000/xyz123.')).toBe('https://doi.org/10.1000/xyz123');
    expect(linkUrl('https://dx.doi.org/10.1000/abc')).toBe('https://doi.org/10.1000/abc');
  });
  it('accepts web and mail addresses', () => {
    expect(linkUrl('https://example.com/a?b=1')).toBe('https://example.com/a?b=1');
    expect(linkUrl('www.apa.org/style')).toBe('https://www.apa.org/style');
    expect(linkUrl('apastyle.apa.org')).toBe('https://apastyle.apa.org');
    expect(linkUrl('me@school.edu')).toBe('mailto:me@school.edu');
  });
  it('refuses words', () => {
    expect(linkUrl('')).toBeNull();
    expect(linkUrl('two words')).toBeNull();
    expect(linkUrl('hello')).toBeNull();
  });
});

describe('links on screen', () => {
  const p = para(0, [{t: 'See the '}, {t: 'APA site', b: true}, {t: ' for more.'}]);

  it('links and unlinks a range, text unchanged', () => {
    const linked = applyOps([p], [{op: 'link', para: 0, start: 4, end: 16, url: 'https://apa.org'}])[0] as ParagraphBlock;
    expect(paragraphText(linked)).toBe(paragraphText(p));
    expect(linked.runs.filter(r => r.l).map(r => r.t).join('')).toBe('the APA site');
    // A caret inside the link removes all of it, across its runs.
    expect(linkSpan(linked, 10, 10)).toEqual({start: 4, end: 16});
    const un = applyOps([linked], [{op: 'unlink', para: 0, start: 10, end: 10}])[0] as ParagraphBlock;
    expect(un.runs.some(r => r.l)).toBe(false);
    expect(un.runs.find(r => r.t.includes('APA'))?.b).toBe(true);
  });

  it('refuses links over links, fields and objects', () => {
    const linked = applyOps([p], [{op: 'link', para: 0, start: 4, end: 16, url: 'https://apa.org'}])[0] as ParagraphBlock;
    expect(linkProblem(linked, 0, 6)).toMatch(/already a link/);
    expect(linkProblem(p, 3, 3)).toMatch(/Select/);
    expect(linkProblem(para(0, [{t: 'a'}, {t: OBJECT, obj: 'note'}, {t: 'b'}]), 0, 3)).toMatch(/field, note or picture/);
    expect(linkProblem(p, 0, 3)).toBeNull();
    expect(linkSpan(p, 2, 2)).toBeNull();
  });
});

describe('header and page numbers', () => {
  it('the last header op wins; logos stay flagged', () => {
    const h = headerFooterAfter({text: '', pageNumber: false, align: 'left', other: true}, 'header', [
      {op: 'headerFooter', para: -1, kind: 'header', text: 'A', pageNumber: false, align: 'right'},
      {op: 'headerFooter', para: -1, kind: 'footer', text: '', pageNumber: true, align: 'center'},
      {op: 'headerFooter', para: -1, kind: 'header', text: 'B', pageNumber: true, align: 'right'},
    ]);
    expect(h).toEqual({text: 'B', pageNumber: true, align: 'right', other: true});
  });

  it('header ops leave the text alone', () => {
    const blocks: Block[] = [para(0, [{t: 'Hello'}])];
    expect(applyOps(blocks, [{op: 'headerFooter', para: -1, kind: 'header', text: 'X', pageNumber: true, align: 'right'}])).toEqual(blocks);
  });

  it('presets put the page number top right, MLA after the last name', () => {
    const blocks: Block[] = [para(0, [{t: 'Text.'}])];
    const hf = (ops: ReturnType<typeof presetOps>) => ops.find(o => o.op === 'headerFooter');
    expect(hf(presetOps(blocks, 'apa', 'Reatherford'))).toMatchObject({kind: 'header', text: '', pageNumber: true, align: 'right'});
    expect(hf(presetOps(blocks, 'mla', ' Reatherford '))).toMatchObject({text: 'Reatherford', pageNumber: true});
    expect(hf(presetOps(blocks, 'chicago'))).toMatchObject({text: '', pageNumber: true});
  });
});
