import {applyOps} from '../src/domain/edits';
import {quoteWithReference} from '../src/domain/quotes';
import {formatSource} from '../src/domain/reference';
import {paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const src = formatSource({type: 'article', authors: [{family: 'Ray', given: 'Oakley'}], year: '2004', title: 'How the Mind Hurts', container: 'American Psychologist', volume: '59', issue: '1', pages: '29-40'}, 'apa');
const p = (index: number, t: string): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t}]});

describe('inserting a quote', () => {
  it('short: in the sentence, quoted and cited, plus the reference', () => {
    const blocks: Block[] = [p(0, 'Stress matters.')];
    const r = quoteWithReference(blocks, applyOps, {para: 0, offset: 14}, 'the mind can hurt the body', src, 'apa', '31');
    const out = applyOps(blocks, r.ops) as ParagraphBlock[];
    expect(paragraphText(out[0])).toBe('Stress matters “the mind can hurt the body” (Ray, 2004, p. 31).');
    expect(out.map(paragraphText)[1]).toBe('References');
    expect(r.block).toBe(false);
  });

  it('40 words or more: a block quotation after the paragraph', () => {
    const long = Array.from({length: 42}, (_, i) => `w${i}`).join(' ');
    const blocks: Block[] = [p(0, 'As Ray argued:'), p(1, 'Next.')];
    const r = quoteWithReference(blocks, applyOps, {para: 0, offset: 14}, long, src, 'apa', '31');
    const out = applyOps(blocks, r.ops) as ParagraphBlock[];
    expect(r.block).toBe(true);
    expect(out[1]).toMatchObject({quote: true});
    expect(paragraphText(out[1])).toBe(`${long}. (Ray, 2004, p. 31)`);
    expect(paragraphText(out[2])).toBe('Next.');
  });
});
