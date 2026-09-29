import {applyOps} from '../src/domain/edits';
import {listFormat, listKind, newListDef, recount} from '../src/domain/lists';
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

import {countWords} from '../src/model/docx';

describe('countWords', () => {
  it('counts words and characters, not objects or stray punctuation', () => {
    expect(countWords(['Hello, world — again.', '￼ image', ''])).toEqual({words: 4, chars: 23});
  });
});

describe('list kinds and levels (2026-09-29)', () => {
  const item = (index: number, id: string, lvl = 0): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 720, runs: [{t: 'x'}], num: {id, lvl}, list: ''});
  const labels = (bs: Block[]) => bs.map(b => (b.type === 'p' ? b.list : ''));

  it('numbers each kind as Word does, level by level', () => {
    expect(labels(recount([item(0, 'l1'), item(1, 'l1')], {}))).toEqual(['a.', 'b.']);
    expect(labels(recount([item(0, 'u1'), item(1, 'u1')], {}))).toEqual(['A.', 'B.']);
    expect(labels(recount([item(0, 'r1'), item(1, 'r1'), item(2, 'r1'), item(3, 'r1')], {}))).toEqual(['i.', 'ii.', 'iii.', 'iv.']);
    expect(labels(recount([item(0, 'o1'), item(1, 'o1', 1), item(2, 'o1', 2), item(3, 'o1', 1), item(4, 'o1')], {}))).toEqual(['I.', 'A.', '1.', 'B.', 'II.']);
    expect(labels(recount([item(0, 'n1'), item(1, 'n1', 1), item(2, 'n1', 2)], {}))).toEqual(['1.', 'a.', 'i.']);
    expect(listKind('o3')).toBe('outline');
    expect(listFormat('outline', 4)).toBe('lowerRoman');
    expect(listFormat('outline', 5)).toBe('decimal');
  });

  it('Tab moves an item a level in, Shift+Tab out, within 0–8', () => {
    const [b] = applyOps([item(0, 'n1')], [{op: 'listLevel', para: 0, delta: 1}]) as ParagraphBlock[];
    expect(b.num).toEqual({id: 'n1', lvl: 1});
    expect(b.indent).toBe(1440);
    const [c] = applyOps([b], [{op: 'listLevel', para: 0, delta: -1}, {op: 'listLevel', para: 0, delta: -1}]) as ParagraphBlock[];
    expect(c.num?.lvl).toBe(0);
    expect(c.indent).toBe(720);
  });
});
