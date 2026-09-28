import {applyOps, fromShown, pageBreakChars, toShown} from '../src/domain/edits';
import {pageInfo, pageMarks} from '../src/domain/pageInfo';
import {pageIndexOf, pageOfChar, type PageStart} from '../src/domain/paging';
import {OBJECT, type Block, type ParagraphBlock} from '../src/model/docx';

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

const page = (block: number, offset = 0, char = 0): PageStart => ({anchor: {block, offset}, char});

describe('the caret drawn in the text', () => {
  const p = para(0, 'Hello world', {caret: 5});

  it('counts as one shown character at its place', () => {
    expect(toShown(p, 5)).toBe(5); // the caret glyph itself
    expect(toShown(p, 6)).toBe(7);
    expect(fromShown(p, 5)).toBe(5); // a pen on the caret: its place
    expect(fromShown(p, 7)).toBe(6);
    expect(fromShown(p, 3)).toBe(3);
  });

  it('comes after deleted text shown at the same place', () => {
    const q = para(0, 'abcdef', {caret: 3, revs: [{id: '1', kind: 'del', author: '', date: '', at: 3, runs: [{t: 'XY'}]}]});
    // a b c X Y | d e f
    expect(fromShown(q, 3)).toBe(3); // on X
    expect(fromShown(q, 5)).toBe(3); // on the caret
    expect(fromShown(q, 6)).toBe(3); // d
    expect(fromShown(q, 7)).toBe(4); // e
  });
});

describe('page cards', () => {
  const blocks: Block[] = [
    para(0, 'Title', {kind: 'heading', level: 1}),
    para(1, 'First page text'),
    para(2, 'Method', {kind: 'heading', level: 3}),
    para(3, `A long paragraph with a note ${OBJECT} and more words`, {
      runs: [{t: 'A long paragraph with a note '}, {t: OBJECT, obj: 'ink', ink: '7'}, {t: ' and more words'}],
      marks: [
        {id: '1', kind: 'ref', at: 2},
        {id: '2', kind: 'ref', at: 2},
        {id: '3', kind: 'ref', at: 40},
      ],
    }),
    para(4, 'Last', {runs: [{t: 'La'}, {t: 'st', rv: '9'}], revs: [{id: '9', kind: 'ins', author: '', date: ''}]}),
  ];
  // Page 2 starts at the Method heading; page 3 starts inside paragraph 3 at character 35.
  const pages = [page(0), page(2), page(3, 40, 35)];
  const replies = (id: string) => id === '2';

  it('shows only a heading that starts on the page', () => {
    expect(pageInfo(blocks, pages, 0, replies).heading).toBe('Title');
    expect(pageInfo(blocks, pages, 1, replies).heading).toBe('Method');
    // Page 3 starts part-way into a body paragraph; no heading starts there (the old cards repeated "Method").
    expect(pageInfo(blocks, pages, 2, replies).heading).toBe('');
  });

  it('counts comments, notes and changes where they sit', () => {
    const p2 = pageInfo(blocks, pages, 1, replies);
    expect(p2).toMatchObject({comments: 1, notes: 1, changes: 0}); // the reply is not counted again
    const p3 = pageInfo(blocks, pages, 2, replies);
    expect(p3).toMatchObject({comments: 1, notes: 0, changes: 1});
    expect(pageInfo(blocks, pages, 0, replies)).toMatchObject({comments: 0, notes: 0, changes: 0});
    expect(pageMarks(p2)).toBe('Comment 1 · Note 1');
    expect(pageMarks(p3)).toBe('Comment 1 · Change 1');
    expect(pageMarks(pageInfo(blocks, pages, 0, replies))).toBe('');
  });

  it('opens with the words at the top of the page', () => {
    expect(pageInfo(blocks, pages, 2, replies).text.startsWith('more words Last')).toBe(true);
  });
});

describe('page numbers', () => {
  it('finds the page a heading is on, for jumps and ◀', () => {
    const pages = [page(0), page(4), page(4, 300, 120), page(9)];
    expect(pageIndexOf(pages, {block: 5, offset: 0})).toBe(2);
    expect(pageIndexOf(pages, {block: 4, offset: 0})).toBe(1);
    expect(pageIndexOf(pages, {block: 9, offset: 0})).toBe(3);
  });
});

describe('the page a character is on', () => {
  // Block 4 runs from page 2 (index 1) onto page 3, which starts at its character 120.
  const pages = [page(0), page(4), page(4, 300, 120), page(9)];
  it('follows a paragraph across a page break', () => {
    expect(pageOfChar(pages, 4, 50, true)).toBe(1);
    expect(pageOfChar(pages, 4, 120, true)).toBe(2);
    expect(pageOfChar(pages, 6, 0, true)).toBe(2);
    expect(pageOfChar(pages, 9, 3, true)).toBe(3);
  });
  it('is unknown past what has been counted', () => {
    expect(pageOfChar(pages, 12, 0, false)).toBe(-1);
    expect(pageOfChar(pages, 6, 0, false)).toBe(2);
  });
});

describe('Word page breaks', () => {
  const brk = para(3, 'end.\nNext', {runs: [{t: 'end.'}, {t: '\n', pg: true}, {t: 'Next'}]});

  it('are found, but shown as plain line breaks (screen pages reflow freely)', () => {
    expect(pageBreakChars(brk)).toEqual([4]);
    expect(toShown(brk, 5)).toBe(5);
    expect(fromShown(brk, 5)).toBe(5);
  });

  it('typed text beside one is plain text, and a split keeps the new page for the first half only', () => {
    const typed = applyOps([brk], [{op: 'text', para: 3, start: 5, end: 5, text: 'X'}])[0] as ParagraphBlock;
    expect(typed.runs.find(r => r.t.includes('X'))?.pg).toBeUndefined();
    const starts = para(0, 'Heading text', {pb: true});
    const [a, b] = applyOps([starts], [{op: 'split', para: 0, offset: 7}]) as ParagraphBlock[];
    expect(a.pb).toBe(true);
    expect(b.pb).toBeUndefined();
  });
});
