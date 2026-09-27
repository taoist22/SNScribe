import {applyOps, pageAfter} from '../src/domain/edits';
import {presetOps} from '../src/domain/presets';
import type {Block, ParagraphBlock} from '../src/model/docx';

const p = (index: number, text: string, kind: ParagraphBlock['kind'] = 'body', level = 0): ParagraphBlock => ({
  type: 'p', index, style: '', kind, level, align: 'left', indent: 0, runs: [{t: text}],
});

const paper: Block[] = [
  p(0, 'My Paper', 'title'),
  p(1, 'Introduction', 'heading', 1),
  p(2, 'Body text here.'),
  p(3, 'References'),
  p(4, 'Smith, J. (2020). A book.'),
];

describe('presetOps', () => {
  it('formats an APA paper', () => {
    const out = applyOps(paper, presetOps(paper, 'apa')) as ParagraphBlock[];
    expect(out[0]).toMatchObject({align: 'center', line: 480, first: undefined});
    expect(out[0].runs[0]).toMatchObject({f: 'Times New Roman', sz: 24, b: true});
    expect(out[1]).toMatchObject({align: 'center'});
    expect(out[2]).toMatchObject({line: 480, before: 0, after: 0, first: 720, align: 'left'});
    expect(out[3]).toMatchObject({align: 'center'});
    expect(out[4]).toMatchObject({first: -720, indent: 720, line: 480});
    expect(pageAfter(undefined, presetOps(paper, 'apa'))).toMatchObject({top: 1440, left: 1440});
  });

  it('keeps MLA titles plain and uses Works Cited', () => {
    const mla: Block[] = [p(0, 'Title', 'title'), p(1, 'Text.'), p(2, 'Works Cited'), p(3, 'Entry.')];
    const out = applyOps(mla, presetOps(mla, 'mla')) as ParagraphBlock[];
    expect(out[0].runs[0].b).toBe(false);
    expect(out[3]).toMatchObject({first: -720});
  });

  it('single-spaces a Chicago bibliography with a blank line between', () => {
    const ch: Block[] = [p(0, 'Text.'), p(1, 'Bibliography'), p(2, 'Entry.')];
    const out = applyOps(ch, presetOps(ch, 'chicago')) as ParagraphBlock[];
    expect(out[2]).toMatchObject({line: 240, after: 240, first: -720});
  });
});

describe('document defaults', () => {
  it('changes the base font and size shown for every paragraph', () => {
    const out = applyOps(paper, [{op: 'defaults', para: -1, font: 'Georgia', size: 28}]) as ParagraphBlock[];
    expect(out.every(q => q.bf === 'Georgia' && q.bs === 28)).toBe(true);
  });
});
