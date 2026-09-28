import {applyOps, expectedTexts} from '../src/domain/edits';
import {tableOps} from '../src/domain/figures';
import {paragraphText, type Block, type ParagraphBlock, type TableBlock} from '../src/model/docx';

const para = (index: number, t: string): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: t ? [{t}] : []});
const texts = (blocks: Block[]) => blocks.filter((b): b is ParagraphBlock => b.type === 'p').map(paragraphText);
const tables = (blocks: Block[]) => blocks.filter((b): b is TableBlock => b.type === 'table');

describe('tables', () => {
  const existing: TableBlock = {type: 'table', rows: 1, cols: 2, preview: '', t: 0, widths: [1, 1], grid: [[{p: [{t: 'a'}]}, {p: [{t: 'b'}]}]]};
  const blocks: Block[] = [para(0, 'Intro.'), para(1, 'Table 1'), para(2, 'Old'), existing, para(3, 'After.')];

  it('APA: label and title above; later tables renumbered and re-addressed', () => {
    const {ops, n} = tableOps(blocks, blocks[0] as ParagraphBlock, {rows: 2, cols: 3, header: true}, 'Scores by Group', 'apa');
    expect(n).toBe(1);
    const after = applyOps(blocks, ops);
    expect(texts(after)).toEqual(['Intro.', 'Table 1', 'Scores by Group', '', 'Table 2', 'Old', 'After.']);
    const [made, old] = tables(after);
    expect(made).toMatchObject({t: 0, rows: 2, cols: 3});
    expect(old.t).toBe(1);
    // Right before the empty paragraph that follows it.
    expect(after[after.indexOf(made) + 1]).toMatchObject({type: 'p', index: 3});
    expect(expectedTexts(blocks, ops)).toEqual(texts(after));
    expect(((after[2] as ParagraphBlock).runs[0]).i).toBe(true);
  });

  it('cells, rows added and deleted', () => {
    let b = applyOps(blocks, [{op: 'tableCell', para: -1, table: 0, row: 0, cell: 1, pieces: [{t: 'Mean', b: true}]}]);
    expect(tables(b)[0].grid![0][1].p).toEqual([{t: 'Mean', b: true}]);
    b = applyOps(b, [{op: 'tableRowAdd', para: -1, table: 0, row: 0, below: true}]);
    expect(tables(b)[0].rows).toBe(2);
    expect(tables(b)[0].grid![1].map(c => c.p)).toEqual([[], []]);
    b = applyOps(b, [{op: 'tableRowDelete', para: -1, table: 0, row: 0}]);
    expect(tables(b)[0].grid!.map(r => r.map(c => c.p.length))).toEqual([[0, 0]]);
    // The last row stays.
    b = applyOps(b, [{op: 'tableRowDelete', para: -1, table: 0, row: 0}]);
    expect(tables(b)[0].rows).toBe(1);
  });

  it('Chicago: one label line; no label: just the table', () => {
    const one = [para(0, 'Text.')];
    expect(texts(applyOps(one, tableOps(one, one[0] as ParagraphBlock, {rows: 2, cols: 2, header: false}, 'Data', 'chicago').ops))).toEqual(['Text.', 'Table 1. Data', '']);
    const plain = applyOps(one, tableOps(one, one[0] as ParagraphBlock, {rows: 2, cols: 2, header: false}, null, 'apa').ops);
    expect(plain.map(b => b.type)).toEqual(['p', 'table', 'p']);
  });
});
