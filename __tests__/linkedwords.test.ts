import {applyOps, expectedTexts} from '../src/domain/edits';
import {linkKind, linkSegments, quoteLinkName, quoteOfLink} from '../src/domain/links';
import {quoteWithReference} from '../src/domain/quotes';
import type {Source} from '../src/domain/citations';
import {paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const para = (index: number, t: string, extra: Partial<ParagraphBlock> = {}): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t}], ...extra});

describe('linked words', () => {
  it('finds comment, note and quote ranges, across paragraphs', () => {
    const blocks: Block[] = [
      para(0, 'Alpha beta gamma', {marks: [{id: '3', kind: 'start', at: 6}, {id: '_sns_ink_ab', kind: 'start', at: 0}, {id: '_sns_ink_ab', kind: 'end', at: 5}]}),
      para(1, 'Delta epsilon', {marks: [{id: '3', kind: 'end', at: 5}, {id: '3', kind: 'ref', at: 5}, {id: '9', kind: 'start', at: 0}, {id: '9', kind: 'end', at: 3}]}),
    ];
    const segs = linkSegments(blocks, new Set(['3']));
    expect(segs.get(0)).toEqual([
      {id: '_sns_ink_ab', kind: 'ink', start: 0, end: 5},
      {id: '3', kind: 'comment', start: 6, end: 16},
    ]);
    // Comment 3 ends in paragraph 1; "9" is not a comment thread here, so it is not a link.
    expect(segs.get(1)).toEqual([{id: '3', kind: 'comment', start: 0, end: 5}]);
    expect(linkKind('_sns_q_x_1', new Set())).toBe('quote');
  });

  it('names a quote link after its saved quote', () => {
    const name = quoteLinkName('q-1759000000-abc');
    expect(name.length).toBeLessThanOrEqual(40);
    expect(quoteOfLink(name, [{id: 'other'}, {id: 'q-1759000000-abc'}])?.id).toBe('q-1759000000-abc');
  });

  it('a quote reports where its words are, and a bookmark goes around them', () => {
    const source: Source = {key: 'k', title: 'T', authors: 'Smith', year: '2020', citation: '<span>(Smith, 2020)</span>', bibHtml: '<div>Smith. T.</div>'};
    const blocks: Block[] = [para(0, 'As she wrote, more text.')];
    const r = quoteWithReference(blocks, applyOps, {para: 0, offset: 13}, 'it was so', source, 'apa', '4');
    const after = applyOps(blocks, r.ops);
    const text = paragraphText(after[0] as ParagraphBlock);
    expect(text.slice(r.span.start, r.span.end)).toBe('“it was so”');
    const ops = [...r.ops, {op: 'bookmark' as const, para: -1 as const, fromPara: 0, from: r.span.start, toPara: 0, to: r.span.end, name: '_sns_q_x_1'}];
    const linked = applyOps(blocks, ops);
    expect(linkSegments(linked, new Set()).get(0)).toEqual([{id: '_sns_q_x_1', kind: 'quote', start: r.span.start, end: r.span.end}]);
    expect(expectedTexts(blocks, ops)).toEqual(linked.map(b => (b.type === 'p' ? paragraphText(b) : '')).filter((_, i) => linked[i].type === 'p'));
  });
});
