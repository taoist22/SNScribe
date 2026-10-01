import {reviewItems} from '../src/domain/reviewlist';
import type {Block, Comment, ParagraphBlock} from '../src/model/docx';

const para = (index: number, extra: Partial<ParagraphBlock>): ParagraphBlock => ({
  type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t: ''}], ...extra,
});

const blocks: Block[] = [
  // "The thesis is weak here." — comment 5 on "thesis is weak" (4..18).
  para(0, {
    runs: [{t: 'The thesis is weak here.'}],
    marks: [{id: '5', kind: 'start', at: 4}, {id: '5', kind: 'end', at: 18}, {id: '5', kind: 'ref', at: 18}],
  }),
  // A note (✎ at 0) about "Second" (bookmark 1..7), then "big " inserted (id 7).
  para(1, {
    runs: [{t: '￼', obj: 'ink', ink: 'n1'}, {t: 'Second '}, {t: 'big ', rv: '7'}, {t: 'point.'}],
    marks: [{id: '_sns_ink_n1', kind: 'start', at: 1}, {id: '_sns_ink_n1', kind: 'end', at: 7}],
    revs: [
      {id: '7', kind: 'ins', author: 'Ann', date: '2026-09-01'},
      {id: '8', kind: 'del', author: 'Ann', date: '2026-09-01', at: 18, runs: [{t: 'very '}]},
    ],
  }),
  // A comment running into the next paragraph.
  para(2, {runs: [{t: 'Start of it'}], marks: [{id: '9', kind: 'start', at: 6}]}),
  para(3, {runs: [{t: 'and the end. More.'}], marks: [{id: '9', kind: 'end', at: 12}]}),
];

const comments: Comment[] = [
  {id: '5', author: 'Prof', initials: 'P', date: '', text: 'Make a claim.'},
  {id: '6', author: 'Me', initials: 'M', date: '', text: 'Fixed', parent: '5'},
  {id: '9', author: 'Prof', initials: 'P', date: '', text: 'Long range'},
];

describe('the review list', () => {
  const items = reviewItems(blocks, comments);

  it('lists comments, notes and changes in reading order', () => {
    expect(items.map(r => `${r.kind}:${r.id}`)).toEqual(['comment:5', 'ink:_sns_ink_n1', 'change:1|7', 'change:1|8', 'comment:9']);
  });

  it('gives a comment its words, text and reply count', () => {
    expect(items[0]).toMatchObject({para: 0, offset: 4, words: 'thesis is weak', text: 'Make a claim.', author: 'Prof', replies: 1});
  });

  it('gives a note the words it is about', () => {
    expect(items[1]).toMatchObject({para: 1, offset: 0, words: 'Second'});
  });

  it('gives changes their text and where they are', () => {
    expect(items[2]).toMatchObject({offset: 8, text: 'big', del: false, author: 'Ann'});
    expect(items[3]).toMatchObject({offset: 18, text: 'very', del: true});
  });

  it('follows a comment range into the next paragraph', () => {
    expect(items[4].words).toBe('of it and the end.');
  });

  it('is empty for a plain document', () => {
    expect(reviewItems([para(0, {runs: [{t: 'Plain.'}]})], [])).toEqual([]);
  });
});
