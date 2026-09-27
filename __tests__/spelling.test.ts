import {misspelledRanges, tokens} from '../src/domain/spelling';
import type {ParagraphBlock} from '../src/model/docx';

const words = (s: string) => tokens(s).map(t => t.word);

describe('what counts as a word', () => {
  it('checks words, skips what is not', () => {
    expect(words("Don't recieve the 2nd H2O dose; see https://doi.org/10.1037/abc or APA rules.")).toEqual(["Don't", 'recieve', 'the', 'dose', 'see', 'or', 'rules']);
    expect(words('Email me@school.edu about notes.pdf')).toEqual(['Email', 'about']);
    expect(words('O’Brien’s children’ work')).toEqual(['O’Brien’s', 'children', 'work']);
  });
});

describe('what gets marked', () => {
  const p: ParagraphBlock = {type: 'p', index: 0, style: '', kind: 'body', level: 0, align: 'left', indent: 0, runs: [{t: 'We recieve teh Wilsonian data.'}]};
  const bad = (w: string) => ['recieve', 'teh', 'Wilsonian'].includes(w);

  it('marks misspellings, not added or ignored words, nor the word being typed', () => {
    expect(misspelledRanges(p, bad, new Set(), new Set())).toEqual([{start: 3, end: 10}, {start: 11, end: 14}, {start: 15, end: 24}]);
    expect(misspelledRanges(p, bad, new Set(['wilsonian']), new Set(['teh']))).toEqual([{start: 3, end: 10}]);
    expect(misspelledRanges(p, bad, new Set(), new Set(), 14)).toEqual([{start: 3, end: 10}, {start: 15, end: 24}]);
  });
});
