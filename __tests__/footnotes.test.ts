import {applyOps, footnotesAfter, nextFootnoteId, numberNotes} from '../src/domain/edits';
import {editPieces, noteText} from '../src/domain/footnotes';
import {OBJECT, paragraphText, type ParagraphBlock} from '../src/model/docx';

const para = (index: number, runs: ParagraphBlock['runs']): ParagraphBlock => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs});

describe('footnotes', () => {
  const chicago = [{t: 'Jane Smith, '}, {t: 'A Book', i: true}, {t: ' (Chicago: Press, 2020), 23.'}];

  it('keep italics where the text was not retyped', () => {
    expect(editPieces(chicago, 'Jane Smith, A Book (Chicago: Press, 2020), 45.')).toEqual([{t: 'Jane Smith, '}, {t: 'A Book', i: true}, {t: ' (Chicago: Press, 2020), 45.'}]);
    expect(editPieces(chicago, 'Jane Smith, A Longer Book (Chicago: Press, 2020), 23.')).toEqual([{t: 'Jane Smith, '}, {t: 'A Longer Book', i: true}, {t: ' (Chicago: Press, 2020), 23.'}]);
    expect(noteText(editPieces(chicago, 'Something else.'))).toBe('Something else.');
  });

  it('are added, numbered in reading order, and deleted', () => {
    const blocks = [para(0, [{t: 'One.'}, {t: OBJECT, obj: 'note', fn: '7'}]), para(1, [{t: 'Two.'}])];
    const id = nextFootnoteId([{id: '7', pieces: []}], blocks);
    expect(id).toBe(8);
    const added = applyOps(blocks, [{op: 'footnote', para: 1, at: 0, id, pieces: [{t: 'A note.'}]}]);
    const numbered = numberNotes(added) as ParagraphBlock[];
    expect(numbered[0].runs[1].nn).toBe('1');
    expect(numbered[1].runs[0]).toMatchObject({obj: 'note', fn: '8', nn: '2'});
    expect(paragraphText(numbered[1])).toBe(`${OBJECT}Two.`);
    const gone = applyOps(added, [{op: 'footnoteDelete', para: -1, id: 7}]) as ParagraphBlock[];
    expect(paragraphText(gone[0])).toBe('One.');
    expect((numberNotes(gone)[1] as ParagraphBlock).runs[0].nn).toBe('1');
    expect(footnotesAfter([{id: '7', pieces: []}], [{op: 'footnote', para: 1, at: 0, id: 8, pieces: [{t: 'x'}]}, {op: 'footnoteDelete', para: -1, id: 7}]).map(f => f.id)).toEqual(['8']);
  });
});
