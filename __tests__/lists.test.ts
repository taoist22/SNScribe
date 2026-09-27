import {applyOps} from '../src/domain/edits';
import {newListDef, recount} from '../src/domain/lists';
import type {Block, ListDef, ParagraphBlock} from '../src/model/docx';

const p = (index: number, text: string, num?: ParagraphBlock['num']): ParagraphBlock => ({
  type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t: text}], num,
});
const labels = (bs: Block[]) => bs.map(b => (b.type === 'p' ? b.list ?? '' : '#'));

describe('recount', () => {
  const doc: Record<string, ListDef> = {
    7: {levels: [{fmt: 'decimal', text: '%1.', start: 1}, {fmt: 'lowerLetter', text: '%1.%2)', start: 1}], starts: {}},
  };

  it('counts per list and level, restarting deeper levels', () => {
    const blocks = [p(0, 'a', {id: 7, lvl: 0}), p(1, 'b', {id: 7, lvl: 1}), p(2, 'c', {id: 7, lvl: 1}), p(3, 'd', {id: 7, lvl: 0}), p(4, 'e', {id: 7, lvl: 1})];
    expect(labels(recount(blocks, doc))).toEqual(['1.', '1.a)', '1.b)', '2.', '2.a)']);
  });

  it('numbers a new list, and bullets a new bullet list', () => {
    const blocks = [p(0, 'x', {id: 'n1', lvl: 0}), p(1, 'y', {id: 'n1', lvl: 0}), p(2, 'z', {id: 'b1', lvl: 0}), p(3, 'plain')];
    expect(labels(recount(blocks, {}))).toEqual(['1.', '2.', '•', '']);
  });

  it('renumbers after an Enter splits a list item', () => {
    const blocks: Block[] = [p(0, 'one', {id: 7, lvl: 0}), p(1, 'two', {id: 7, lvl: 0})];
    const after = recount(applyOps(blocks, [{op: 'split', para: 0, offset: 3}]), doc);
    expect(labels(after)).toEqual(['1.', '2.', '3.']);
  });

  it('turns a paragraph into a list item and back', () => {
    const blocks: Block[] = [p(0, 'first'), p(1, 'second')];
    const listed = applyOps(blocks, [{op: 'list', para: 0, kind: 'number', listId: 'n1'}, {op: 'list', para: 1, kind: 'number', listId: 'n1'}]);
    expect(labels(recount(listed, {}))).toEqual(['1.', '2.']);
    const out = applyOps(listed, [{op: 'list', para: 0, kind: 'none', listId: ''}]);
    expect(labels(recount(out, {}))).toEqual(['', '1.']);
  });

  it('matches the writer: a new list starts at 1', () => {
    expect(newListDef('number').starts[0]).toBe(1);
  });
});
