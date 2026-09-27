import {parseNotes, placeNote, reanchor, type InkNote} from '../src/domain/notes';
import type {Block} from '../src/model/docx';

const p = (index: number, t: string): Block => ({type: 'p', index, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t}]});
const note = (para: number, at: number, quote: string): InkNote => ({id: 'n', path: '/x.png', width: 10, height: 5, para, at, quote, created: ''});

describe('placing handwritten notes', () => {
  const doc = [p(0, 'Intro text.'), p(1, 'The method was new.'), p(2, 'Another method here.')];

  it('stays where its words are', () => {
    expect(placeNote(note(1, 4, 'method'), doc)).toEqual({para: 1, at: 4});
  });

  it('follows its words when text moves', () => {
    const edited = [p(0, 'Intro text.'), p(1, 'Now the method was new.'), p(2, 'Another method here.')];
    expect(placeNote(note(1, 4, 'method'), edited)).toEqual({para: 1, at: 8});
    // The paragraph was split off above: the nearest occurrence wins.
    const moved = [p(0, 'Intro text.'), p(1, 'Added.'), p(2, 'The method was new.'), p(3, 'Another method here.')];
    expect(placeNote(note(1, 4, 'method'), moved)).toEqual({para: 2, at: 4});
  });

  it('is unplaced when its words are gone, and kept', () => {
    const gone = [p(0, 'Intro text.')];
    expect(placeNote(note(1, 4, 'method'), gone)).toBeNull();
    const n = note(1, 4, 'method');
    expect(reanchor([n], gone)).toEqual([n]);
  });

  it('an empty-line note keeps its paragraph', () => {
    expect(placeNote(note(2, 50, ''), doc)).toEqual({para: 2, at: 20});
    expect(placeNote(note(9, 0, ''), doc)).toBeNull();
  });

  it('reads stored notes defensively', () => {
    expect(parseNotes('nonsense')).toEqual([]);
    expect(parseNotes(JSON.stringify([note(0, 0, ''), {bad: 1}]))).toHaveLength(1);
  });
});
