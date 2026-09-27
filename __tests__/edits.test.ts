import {
  allHave,
  applyOps,
  deletionRange,
  expectedTexts,
  joinProblem,
  splitProblem,
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
    expect(expectedTexts(doc, ops)).toEqual(['A Title', 'Simple then bold note' + OBJECT + ' end.']);
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

describe('paragraphs', () => {
  const three: Block[] = [
    para(0, [{t: 'Heading'}], 'heading'),
    para(1, [{t: 'First '}, {t: 'bold', b: true}, {t: ' end.'}]),
    para(2, [{t: 'Second.'}]),
  ];
  const texts = (bs: Block[]) => bs.filter(b => b.type === 'p').map(b => paragraphText(b as ParagraphBlock));
  const indexes = (bs: Block[]) => bs.filter(b => b.type === 'p').map(b => (b as ParagraphBlock).index);

  it('splits, keeping run formatting, and renumbers what follows', () => {
    const out = applyOps(three, [{op: 'split', para: 1, offset: 8}]); // "First bo" | "ld end."
    expect(texts(out)).toEqual(['Heading', 'First bo', 'ld end.', 'Second.']);
    expect(indexes(out)).toEqual([0, 1, 2, 3]);
    expect((out[2] as ParagraphBlock).runs[0]).toMatchObject({t: 'ld', b: true});
  });

  it('gives a body paragraph after Enter at the end of a heading', () => {
    const out = applyOps(three, [{op: 'split', para: 0, offset: 7}]);
    expect(out[1]).toMatchObject({kind: 'body', index: 1, runs: []});
    expect(out[0]).toMatchObject({kind: 'heading'});
  });

  it('joins onto the previous paragraph and renumbers', () => {
    const out = applyOps(three, [{op: 'join', para: 2}]);
    expect(texts(out)).toEqual(['Heading', 'First bold end.Second.']);
    expect(indexes(out)).toEqual([0, 1]);
  });

  it('round-trips split then join', () => {
    expect(texts(applyOps(three, [{op: 'split', para: 1, offset: 3}, {op: 'join', para: 2}]))).toEqual(texts(three));
  });

  it('refuses joins across a table or a section break, and splits inside fields', () => {
    expect(joinProblem(doc, 1)).toMatch(/table/);
    expect(joinProblem(three, 0)).toMatch(/no paragraph before/);
    const sect: Block[] = [{...(three[1] as ParagraphBlock), sect: true}, {...(three[2] as ParagraphBlock)}];
    expect(joinProblem(sect, 2)).toMatch(/section break/);
    const field = para(5, [{t: 'Page '}, {t: '12', k: true}, {t: '3', k: true}]);
    expect(splitProblem(field, 6)).toMatch(/field/);
    expect(splitProblem(field, 7)).toMatch(/field/);
    expect(splitProblem(field, 5)).toBeNull();
  });
});

describe('fonts and sizes', () => {
  it('sets font and size on exactly the range', () => {
    const [, , out] = applyOps(doc, [{op: 'runStyle', para: 1, start: 6, end: 15, font: 'Georgia', size: 28}]) as ParagraphBlock[];
    const styled = out.runs.filter(r => r.f === 'Georgia');
    expect(styled.map(r => r.t).join('')).toBe('then bold');
    expect(styled.every(r => r.sz === 28)).toBe(true);
    expect(out.runs.find(r => r.t === 'bold')?.b).toBe(true);
    expect(paragraphText(out)).toBe(paragraphText(p1));
  });
  it('can change only the size', () => {
    const [, , out] = applyOps(doc, [{op: 'runStyle', para: 1, start: 0, end: 5, size: 40}]) as ParagraphBlock[];
    expect(out.runs[0]).toMatchObject({t: 'Plain', sz: 40});
    expect(out.runs[0].f).toBeUndefined();
  });
});

describe('paragraph formatting', () => {
  const one = (ops: Op[]) => applyOps(doc, ops)[2] as ParagraphBlock;
  it('sets alignment, spacing and a first-line indent without touching text', () => {
    const p = one([{op: 'para', para: 1, align: 'center', line: 480, before: 240, after: 0, first: 720}]);
    expect(p).toMatchObject({align: 'center', line: 480, lineRule: 'auto', before: 240, after: 0, first: 720});
    expect(paragraphText(p)).toBe(paragraphText(p1));
  });
  it('gives a hanging indent its left indent, and takes both back', () => {
    const hung = one([{op: 'para', para: 1, first: -720}]);
    expect(hung).toMatchObject({first: -720, indent: 720});
    const back = applyOps([doc[0], doc[1], hung], [{op: 'para', para: 1, first: 0}])[2] as ParagraphBlock;
    expect(back.first).toBeUndefined();
    expect(back.indent).toBe(0);
  });
  it('turns a page break before on and off', () => {
    expect(one([{op: 'para', para: 1, pb: true}]).pb).toBe(true);
    expect(one([{op: 'para', para: 1, pb: true}, {op: 'para', para: 1, pb: false}]).pb).toBeUndefined();
  });
});
