import {applyOps, findMatches, formatOps, highlightOps, replaceAllOps} from '../src/domain/edits';
import {presetOps} from '../src/domain/presets';
import {OBJECT, paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const p = (index: number, runs: ParagraphBlock['runs']): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs});

describe('highlight colours', () => {
  it('sets and removes a colour', () => {
    const [a] = applyOps([p(0, [{t: 'Hello world'}])], highlightOps([{para: 0, start: 0, end: 5}], 'green')) as ParagraphBlock[];
    expect(a.runs[0]).toMatchObject({t: 'Hello', h: true, hc: 'green'});
    const [b] = applyOps([a], highlightOps([{para: 0, start: 0, end: 5}], null)) as ParagraphBlock[];
    expect(b.runs[0].h).toBe(false);
    expect(b.runs[0].hc).toBeUndefined();
  });
});

describe('strike and scripts', () => {
  it('superscript replaces subscript', () => {
    const blocks: Block[] = [p(0, [{t: 'H2O'}])];
    const sub = applyOps(blocks, formatOps(blocks, [{para: 0, start: 1, end: 2}], 'sub'));
    const sup = applyOps(sub, formatOps(sub, [{para: 0, start: 1, end: 2}], 'sup')) as ParagraphBlock[];
    expect(sup[0].runs[1]).toMatchObject({t: '2', sup: true, sub: false});
    const s = applyOps(blocks, formatOps(blocks, [{para: 0, start: 0, end: 3}], 's')) as ParagraphBlock[];
    expect(s[0].runs[0].s).toBe(true);
  });
});

describe('Heading 3 and Quote', () => {
  it('quote stays body text, flagged', () => {
    const [q] = applyOps([p(0, [{t: 'Said.'}])], [{op: 'style', para: 0, kind: 'quote'}]) as ParagraphBlock[];
    expect(q).toMatchObject({kind: 'body', quote: true});
    const [h] = applyOps([q], [{op: 'style', para: 0, kind: 'heading3'}]) as ParagraphBlock[];
    expect(h).toMatchObject({kind: 'heading', level: 3, quote: undefined});
  });

  it('presets keep a quote as a block (no first-line indent), not a heading', () => {
    const blocks: Block[] = [{...p(0, [{t: 'A long quotation.'}]), quote: true}];
    const ops = presetOps(blocks, 'apa');
    expect(ops.find(o => o.op === 'para')).toMatchObject({first: 0});
    expect(ops.some(o => o.op === 'format' && o.prop === 'b')).toBe(false);
  });
});

describe('find & replace', () => {
  const blocks: Block[] = [p(0, [{t: 'The cat and the Cat.'}]), p(1, [{t: 'cat'}, {t: OBJECT, obj: 'image'}, {t: 'cat', k: true}])];

  it('finds with and without case', () => {
    expect(findMatches(blocks, 'cat', false)).toHaveLength(4);
    expect(findMatches(blocks, 'Cat', true)).toEqual([{para: 0, start: 16, end: 19}]);
    expect(findMatches(blocks, '', false)).toEqual([]);
  });

  it('replaces all, later matches first, skipping fields', () => {
    const {ops, skipped} = replaceAllOps(blocks, findMatches(blocks, 'cat', false), 'dog');
    expect(skipped).toBe(1);
    const out = applyOps(blocks, ops) as ParagraphBlock[];
    expect(paragraphText(out[0])).toBe('The dog and the dog.');
    expect(paragraphText(out[1])).toBe(`dog${OBJECT}cat`);
  });
});

import {penSelection} from '../src/domain/edits';

describe('pen selection', () => {
  const t = 'Water is H2O here.';
  const at = (char: number) => ({para: 0, char, offset: char});

  it('inside one word: exactly the characters crossed', () => {
    expect(penSelection(t, t, at(10), at(10))).toEqual({from: {para: 0, offset: 10}, to: {para: 0, offset: 11}}); // "2"
    expect(penSelection(t, t, at(9), at(10))).toEqual({from: {para: 0, offset: 9}, to: {para: 0, offset: 11}}); // "H2"
  });

  it('across words: whole words', () => {
    expect(penSelection(t, t, at(2), at(10))).toEqual({from: {para: 0, offset: 0}, to: {para: 0, offset: 12}}); // "Water is H2O"
  });
});
