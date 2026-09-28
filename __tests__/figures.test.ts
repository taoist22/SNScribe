import {applyOps, expectedTexts} from '../src/domain/edits';
import {figureLabels, figureOps, pictureAtWidth, pictureEmu} from '../src/domain/figures';
import {OBJECT, paragraphText, type Block, type ParagraphBlock} from '../src/model/docx';

const para = (index: number, t: string, extra: Partial<ParagraphBlock> = {}): ParagraphBlock => ({
  type: 'p',
  index,
  style: '',
  kind: 'body',
  level: 0,
  align: 'left',
  indent: 0,
  runs: t ? [{t}] : [],
  ...extra,
});

const texts = (blocks: Block[]) => blocks.map(b => (b.type === 'p' ? paragraphText(b) : ''));
const pic = {path: '/storage/emulated/0/chart.png', cx: 100, cy: 50, alt: 'Chart'};

describe('picture size', () => {
  it('keeps small pictures at 96 dpi and shrinks big ones to fit', () => {
    expect(pictureEmu(96, 48, 1e9, 1e9)).toEqual({cx: 914400, cy: 457200});
    expect(pictureEmu(2000, 1000, 5943600, 1e9)).toEqual({cx: 5943600, cy: 2971800});
  });
});

describe('figures', () => {
  const blocks: Block[] = [para(0, 'Intro text.'), para(1, 'Figure 1'), para(2, 'Old Figure'), para(3, OBJECT, {runs: [{t: OBJECT, obj: 'image'}]}), para(4, 'End.')];

  it('APA: label and italic title above the picture; later figures renumbered', () => {
    const {ops, n} = figureOps(blocks, blocks[0] as ParagraphBlock, pic, 'Growth Over Time', 'apa');
    expect(n).toBe(1);
    const after = applyOps(blocks, ops);
    expect(texts(after).slice(0, 5)).toEqual(['Intro text.', 'Figure 1', 'Growth Over Time', OBJECT, 'Figure 2']);
    const title = after[2] as ParagraphBlock;
    expect(title.runs[0].i).toBe(true);
    expect((after[1] as ParagraphBlock).runs[0].b).toBe(true);
    expect((after[3] as ParagraphBlock).runs[0]).toMatchObject({obj: 'image', src: pic.path, cx: 100, cy: 50});
    // What the screen expects is what the native save must produce.
    expect(expectedTexts(blocks, ops)).toEqual(texts(after));
    expect(figureLabels(after, 'apa').map(l => l.para)).toEqual([1, 4]);
  });

  it('MLA and Chicago: caption below', () => {
    const one = [para(0, 'Text.')];
    expect(texts(applyOps(one, figureOps(one, one[0] as ParagraphBlock, pic, 'A Chart', 'mla').ops))).toEqual(['Text.', OBJECT, 'Fig. 1. A Chart']);
    expect(texts(applyOps(one, figureOps(one, one[0] as ParagraphBlock, pic, 'A Chart', 'chicago').ops))).toEqual(['Text.', OBJECT, 'Figure 1. A Chart']);
  });

  it('a picture without a label', () => {
    const one = [para(0, 'Text.')];
    const {ops, n} = figureOps(one, one[0] as ParagraphBlock, pic, null, 'apa');
    expect(n).toBe(0);
    expect(texts(applyOps(one, ops))).toEqual(['Text.', OBJECT]);
  });
});

describe('picture sizes', () => {
  it('a share of the text width, proportions kept, never taller than allowed', () => {
    expect(pictureAtWidth(2000, 1000, 0.5, 5943600, 1e9)).toEqual({cx: 2971800, cy: 1485900});
    // A tall picture at full width is held to the height limit, narrower.
    expect(pictureAtWidth(1000, 3000, 1, 5943600, 6000000)).toEqual({cx: 2000000, cy: 6000000});
  });

  it('resizes and deletes a picture in the text', () => {
    const p: Block[] = [para(0, '', {runs: [{t: 'A '}, {t: OBJECT, obj: 'image', src: '/x.png', cx: 10, cy: 5}, {t: ' B'}]})];
    const big = applyOps(p, [{op: 'imageSize', para: 0, at: 2, cx: 40, cy: 20}])[0] as ParagraphBlock;
    expect(big.runs[1]).toMatchObject({obj: 'image', cx: 40, cy: 20});
    const gone = applyOps(p, [{op: 'imageDelete', para: 0, at: 2}]);
    expect(texts(gone)).toEqual(['A  B']);
  });
});
