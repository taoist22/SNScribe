import {applyOps, commentsAfter, expectedTexts, fromShown, nextCommentId, threadIds, toShown} from '../src/domain/edits';
import {paragraphText, type ParagraphBlock} from '../src/model/docx';

// "Hello big world" with "old " deleted before "world" (at 10) and "big " inserted (id 7).
const p: ParagraphBlock = {
  type: 'p', index: 0, style: '', kind: 'body', level: 0, align: 'left', indent: 0,
  runs: [{t: 'Hello '}, {t: 'big ', rv: '7'}, {t: 'world'}],
  revs: [
    {id: '7', kind: 'ins', author: 'T', date: ''},
    {id: '8', kind: 'del', author: 'T', date: '', at: 10, runs: [{t: 'old '}]},
  ],
};

describe('screen offsets with deleted text shown', () => {
  it('maps text offsets past the deletion', () => {
    expect(toShown(p, 5)).toBe(5);
    expect(toShown(p, 10)).toBe(10); // the caret sits before the deleted text
    expect(toShown(p, 11)).toBe(15);
    expect(fromShown(p, 15)).toBe(11);
    expect(fromShown(p, 12)).toBe(10); // inside "old " snaps to where it sits
    expect(fromShown(p, 3)).toBe(3);
  });

  it('counts a first-line indent spacer too', () => {
    const q = {...p, first: 720};
    expect(toShown(q, 0)).toBe(1);
    expect(toShown(q, 11)).toBe(16);
    expect(fromShown(q, 0)).toBe(0);
    expect(fromShown(q, 16)).toBe(11);
  });
});

describe('tracked changes follow edits', () => {
  const one = (ops: Parameters<typeof applyOps>[1]) => applyOps([p], ops);

  it('typing before a deletion moves it', () => {
    const [q] = one([{op: 'text', para: 0, start: 0, end: 0, text: 'Oh, '}]) as ParagraphBlock[];
    expect(q.revs?.find(v => v.kind === 'del')?.at).toBe(14);
  });

  it('deleting the inserted text drops its change', () => {
    const [q] = one([{op: 'text', para: 0, start: 6, end: 10, text: ''}]) as ParagraphBlock[];
    expect(paragraphText(q)).toBe('Hello world');
    expect(q.revs?.map(v => v.kind)).toEqual(['del']);
    expect(q.revs?.[0].at).toBe(6);
  });

  it('split and join carry deletions with their text', () => {
    const out = one([{op: 'split', para: 0, offset: 6}]) as ParagraphBlock[];
    expect(out[1].revs?.find(v => v.kind === 'del')?.at).toBe(4);
    const back = applyOps(out, [{op: 'join', para: 1}]) as ParagraphBlock[];
    expect(back[0].revs?.find(v => v.kind === 'del')?.at).toBe(10);
  });

  it('a review takes the file\'s result', () => {
    const ops = [{op: 'revision' as const, para: 0, id: '8', accept: false, result: {0: {runs: [{t: 'Hello big old world'}]}}}];
    const [q] = one(ops) as ParagraphBlock[];
    expect(paragraphText(q)).toBe('Hello big old world');
    expect(q.revs).toBeUndefined();
    expect(expectedTexts([p], ops)).toEqual(['Hello big old world']);
  });
});

describe('comments on screen', () => {
  const q: ParagraphBlock = {type: 'p', index: 0, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t: 'One two three'}]};
  const add = {op: 'comment' as const, para: -1 as const, id: 5, fromPara: 0, from: 4, toPara: 0, to: 7, text: 'Hm', author: 'CT', initials: 'C', date: ''};

  it('adds anchors, replies beside them, and deletes a thread', () => {
    const [a] = applyOps([q], [add]) as ParagraphBlock[];
    expect(a.marks).toEqual([
      {id: '5', kind: 'start', at: 4},
      {id: '5', kind: 'end', at: 7},
      {id: '5', kind: 'ref', at: 7},
    ]);
    const [b] = applyOps([a], [{...add, id: 6, parent: 5}]) as ParagraphBlock[];
    expect(b.marks?.filter(m => m.id === '6').map(m => m.at)).toEqual([4, 7, 7]);
    const comments = commentsAfter([], [add, {...add, id: 6, parent: 5}]);
    expect(threadIds(comments, '5')).toEqual([5, 6]);
    expect(nextCommentId(comments)).toBe(7);
    const [c] = applyOps([b], [{op: 'uncomment', para: -1, ids: [5, 6]}]) as ParagraphBlock[];
    expect(c.marks).toBeUndefined();
  });

  it('anchors follow typing', () => {
    const [a] = applyOps([q], [add, {op: 'text', para: 0, start: 0, end: 0, text: 'Zero '}]) as ParagraphBlock[];
    // Typed at the very start: the range moves with its text.
    expect(a.marks?.map(m => m.at)).toEqual([9, 12, 12]);
    const [b] = applyOps([q], [add, {op: 'text', para: 0, start: 5, end: 6, text: ''}]) as ParagraphBlock[];
    expect(b.marks?.map(m => m.at)).toEqual([4, 6, 6]);
  });
});
