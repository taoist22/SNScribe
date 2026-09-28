import {bodyRange, targetLabel, wordCounts} from '../src/domain/wordcount';
import type {ParagraphBlock} from '../src/model/docx';

const para = (index: number, t: string, extra: Partial<ParagraphBlock> = {}): ParagraphBlock => ({
  type: 'p',
  index,
  style: '',
  kind: 'body',
  level: 0,
  align: 'left',
  indent: 0,
  runs: [{t}],
  ...extra,
});

describe('body text word count', () => {
  it('leaves out an APA title page and the reference list', () => {
    const blocks = [
      para(0, 'A Title Of Four'),
      para(1, 'Student Name'),
      para(2, 'The Real Title', {pb: true}),
      para(3, 'One two three four five.'),
      para(4, 'References', {pb: true}),
      para(5, 'Smith, J. (2020). A book. Publisher.'),
    ];
    expect(bodyRange(blocks)).toEqual({from: 2, to: 4});
    expect(wordCounts(blocks)).toEqual({all: 21, body: 8});
  });

  it('ends a title page at a Word page break too', () => {
    const blocks = [para(0, 'Title'), para(1, '\n', {runs: [{t: '\n', pg: true}]}), para(2, 'Body words here')];
    expect(wordCounts(blocks).body).toBe(3);
  });

  it('counts everything in a paper without a title page or references', () => {
    const blocks = [para(0, 'Name'), para(1, 'Just an essay with words.')];
    expect(wordCounts(blocks)).toEqual({all: 6, body: 6});
  });

  it('labels progress', () => {
    expect(targetLabel(1234, 2500)).toBe('1,234 / 2,500 words');
  });
});
