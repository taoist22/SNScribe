import {changeRanges} from '../src/domain/edits';
import type {ParagraphBlock} from '../src/model/docx';

// "Hello big world" with "big " inserted (id 7) and "old " deleted before "world" (at 10).
const p: ParagraphBlock = {
  type: 'p', index: 0, style: '', kind: 'body', level: 0, align: 'left', indent: 0,
  runs: [{t: 'Hello '}, {t: 'big ', rv: '7'}, {t: 'world'}],
  revs: [
    {id: '7', kind: 'ins', author: 'T', date: ''},
    {id: '8', kind: 'del', author: 'T', date: '', at: 10, runs: [{t: 'old '}]},
  ],
};

describe('where tracked changes show (a tap there opens the review pane)', () => {
  it('gives the inserted and the struck-through text, in screen offsets', () => {
    // Screen: "Hello big old world" — "big " at 6..10, "old " at 10..14.
    expect(changeRanges(p)).toEqual([
      {id: '7', start: 6, end: 10},
      {id: '8', start: 10, end: 14},
    ]);
  });

  it('counts the first-line indent spacer', () => {
    expect(changeRanges({...p, first: 720})).toEqual([
      {id: '7', start: 7, end: 11},
      {id: '8', start: 11, end: 15},
    ]);
  });

  it('joins an insertion split over runs, and puts deletions at one place one after another', () => {
    const q: ParagraphBlock = {
      ...p,
      runs: [{t: 'ab', rv: '1'}, {t: 'cd', rv: '1', b: true}, {t: 'ef'}],
      revs: [
        {id: '1', kind: 'ins', author: 'T', date: ''},
        {id: '2', kind: 'del', author: 'T', date: '', at: 4, runs: [{t: 'XY'}]},
        {id: '3', kind: 'del', author: 'T', date: '', at: 4, runs: [{t: 'Z'}]},
      ],
    };
    expect(changeRanges(q)).toEqual([
      {id: '1', start: 0, end: 4},
      {id: '2', start: 4, end: 6},
      {id: '3', start: 6, end: 7},
    ]);
  });

  it('has nothing for a paragraph without changes', () => {
    expect(changeRanges({...p, runs: [{t: 'plain'}], revs: undefined})).toEqual([]);
  });
});
