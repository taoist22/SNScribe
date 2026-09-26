import {
  allHave,
  applyOps,
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
