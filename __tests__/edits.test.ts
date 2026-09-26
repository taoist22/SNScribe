import {
  allHave,
  applyOps,
  deletionRange,
  expectedTexts,
  shiftPos,
  textEditProblem,
  formatOps,
  markSelection,
  rangesBetween,
  splitRuns,
  styleOps,
  wordAround,
  type Op,
} from '../src/domain/edits';
import {OBJECT, paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const para = (index: number, runs: ParagraphBlock['runs'], kind: ParagraphBlock['kind'] = 'body'): ParagraphBlock => ({
  type: 'p',
  index,
  style: '',
  kind,
  level: 0,
  align: 'left',
  indent: 0,
  runs,
});

const doc: Block[] = [
  para(0, [{t: 'A Title'}], 'title'),
  {type: 'table', rows: 1, cols: 1, preview: ''},
  para(1, [{t: 'Plain then '}, {t: 'bold', b: true}, {t: ' note'}, {t: OBJECT, obj: 'note', sup: true}, {t: ' end.'}]),
];
const p1 = doc[2] as ParagraphBlock;

describe('splitRuns', () => {
  it('cuts text runs and never an object run', () => {
    const runs = splitRuns(p1.runs, [3, 13, 20]);
    expect(runs.map(r => r.t).join('')).toBe(paragraphText(p1));
    expect(runs.filter(r => r.obj)).toHaveLength(1);
    expect(runs.map(r => r.t)).toContain('Pla');
  });
});

describe('applyOps', () => {
  it('formats exactly the range, keeping text and other styles', () => {
    const [, , out] = applyOps(doc, [{op: 'format', para: 1, start: 6, end: 13, prop: 'h', on: true}]) as ParagraphBlock[];
    expect(paragraphText(out)).toBe(paragraphText(p1));
    const lit = out.runs.filter(r => r.h).map(r => r.t).join('');
    expect(lit).toBe('then bo');
    expect(out.runs.find(r => r.t === 'bo')?.b).toBe(true);
  });

  it('applies in order, so a later op wins, and leaves other blocks alone', () => {
    const ops: Op[] = [
      {op: 'format', para: 1, start: 0, end: 5, prop: 'b', on: true},
      {op: 'format', para: 1, start: 0, end: 5, prop: 'b', on: false},
      {op: 'style', para: 0, kind: 'heading2'},
    ];
    const out = applyOps(doc, ops);
    expect((out[2] as ParagraphBlock).runs.some(r => r.t.startsWith('Plain') && r.b)).toBe(false);
    expect(out[0]).toMatchObject({kind: 'heading', level: 2});
    expect(out[1]).toBe(doc[1]);
  });

  it('can highlight across an object', () => {
    const [, , out] = applyOps(doc, [{op: 'format', para: 1, start: 15, end: 22, prop: 'h', on: true}]) as ParagraphBlock[];
    expect(out.runs.find(r => r.obj)?.h).toBe(true);
  });
});

describe('selections', () => {
  it('turns two positions into per-paragraph ranges, in order', () => {
    expect(rangesBetween(doc, {para: 1, offset: 4}, {para: 0, offset: 2})).toEqual([
      {para: 0, start: 2, end: 7},
      {para: 1, start: 0, end: 4},
    ]);
  });

  it('toggles: bold on unless all of the selection is bold', () => {
    const boldOnly = [{para: 1, start: 11, end: 15}];
    expect(allHave(doc, boldOnly, 'b')).toBe(true);
    expect(formatOps(doc, boldOnly, 'b')[0]).toMatchObject({on: false});
    expect(formatOps(doc, [{para: 1, start: 6, end: 15}], 'b')[0]).toMatchObject({on: true});
  });

  it('makes one style op per paragraph', () => {
    expect(styleOps([{para: 1, start: 0, end: 2}, {para: 1, start: 4, end: 6}], 'heading1')).toEqual([
      {op: 'style', para: 1, kind: 'heading1'},
    ]);
  });

  it('snaps to words, stepping off spaces toward the selection', () => {
    const text = 'Plain words then bold';
    const w = (ch: number, dir: 'left' | 'right') => {
      const r = wordAround(text, ch, dir);
      return r && text.slice(r.start, r.end);
    };
    expect(w(2, 'right')).toBe('Plain');
    expect(w(5, 'right')).toBe('words');
    expect(w(5, 'left')).toBe('Plain');
    expect(wordAround('word  ', 5, 'right')).toBeNull();
  });

  it('marks the selected runs for drawing', () => {
    const marked = markSelection(p1.runs, {start: 6, end: 15});
    expect(marked.filter(r => r.sel).map(r => r.t).join('')).toBe('then bold');
  });
});


describe('text edits', () => {
  const edit = (start: number, end: number, text: string) =>
    (applyOps(doc, [{op: 'text', para: 1, start, end, text}])[2] as ParagraphBlock);

  it('replaces inside a formatted run with that formatting', () => {
    const out = edit(11, 15, 'BRAVE'); // "bold" → "BRAVE"
    expect(paragraphText(out)).toBe('Plain then BRAVE note' + OBJECT + ' end.');
    expect(out.runs.find(r => r.t === 'BRAVE')?.b).toBe(true);
  });

  it('inserts with the formatting of the character before', () => {
    const out = edit(15, 15, '!'); // right after "bold"
    expect(out.runs.find(r => r.t === '!')?.b).toBe(true);
    const start = edit(0, 0, '>> ');
    expect(paragraphText(start).startsWith('>> Plain')).toBe(true);
    expect(start.runs[0]).toMatchObject({t: '>> '});
  });

  it('deletes across runs', () => {
    expect(paragraphText(edit(6, 16, ''))).toBe('Plain note' + OBJECT + ' end.');
  });

  it('never copies object-ness into typed text', () => {
    const out = edit(21, 21, 'x'); // right after the note marker
    expect(out.runs.find(r => r.t === 'x')?.obj).toBeUndefined();
  });

  it('refuses ranges that touch objects or fields', () => {
    expect(textEditProblem(p1, 17, 22)).toMatch(/note marker/);
    expect(textEditProblem(p1, 0, 5)).toBeNull();
    const field = {...p1, runs: [{t: 'Page '}, {t: '3', k: true}]};
    expect(textEditProblem(field, 5, 6)).toMatch(/field/);
    expect(textEditProblem(field, 6, 6)).toMatch(/field/);
  });

  it('removes the extra space when deleting a word', () => {
    const t = 'one two three.';
    expect(deletionRange(t, 4, 7)).toEqual({start: 4, end: 8}); // "two " → "one three."
    expect(deletionRange(t, 8, 13)).toEqual({start: 7, end: 13}); // " three" before "."
    expect(deletionRange(t, 0, 3)).toEqual({start: 0, end: 4});
  });

  it('computes the texts to cross-check on save', () => {
    const ops: Op[] = [{op: 'text', para: 1, start: 0, end: 5, text: 'Simple'}];
    expect(expectedTexts(doc, ops)).toEqual({1: 'Simple then bold note' + OBJECT + ' end.'});
  });
});

describe('shiftPos', () => {
  const r = {para: 2, start: 10, end: 14}; // 4 chars replaced by 7
  it('moves positions after the edit, keeps those before or elsewhere', () => {
    expect(shiftPos({para: 2, offset: 20}, r, 7)).toEqual({para: 2, offset: 23});
    expect(shiftPos({para: 2, offset: 10}, r, 7)).toEqual({para: 2, offset: 10});
    expect(shiftPos({para: 3, offset: 20}, r, 7)).toEqual({para: 3, offset: 20});
  });
  it('puts positions inside the replaced text at its end', () => {
    expect(shiftPos({para: 2, offset: 12}, r, 7)).toEqual({para: 2, offset: 17});
  });
  it('handles a plain insertion', () => {
    expect(shiftPos({para: 2, offset: 11}, {para: 2, start: 10, end: 10}, 3)).toEqual({para: 2, offset: 14});
  });
});
