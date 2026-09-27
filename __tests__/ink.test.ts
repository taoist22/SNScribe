import {applyOps, expectedTexts, inksAfter} from '../src/domain/edits';
import {OBJECT, paragraphText, type ParagraphBlock} from '../src/model/docx';

const p: ParagraphBlock = {type: 'p', index: 0, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t: 'Hello '}, {t: 'world', b: true}]};
const add = {op: 'ink' as const, para: 0, at: 6, id: 'k1', png: '/n/k1.png', width: 400, height: 200};

describe('handwritten margin notes', () => {
  it('anchor as one object character, as the file reads them', () => {
    const [q] = applyOps([p], [add]) as ParagraphBlock[];
    expect(paragraphText(q)).toBe(`Hello ${OBJECT}world`);
    expect(q.runs.find(r => r.obj === 'ink')?.ink).toBe('k1');
    expect(expectedTexts([p], [add])).toEqual([`Hello ${OBJECT}world`]);
    expect(inksAfter({old: '/x.png'}, [add])).toEqual({old: '/x.png', k1: '/n/k1.png'});
  });

  it('delete takes the anchor out again', () => {
    const [q] = applyOps([p], [add, {op: 'inkDelete', para: 0, id: 'k1'}]) as ParagraphBlock[];
    expect(paragraphText(q)).toBe('Hello world');
  });

  it('at the end of a paragraph', () => {
    const [q] = applyOps([p], [{...add, at: 11}]) as ParagraphBlock[];
    expect(paragraphText(q)).toBe(`Hello world${OBJECT}`);
  });
});
