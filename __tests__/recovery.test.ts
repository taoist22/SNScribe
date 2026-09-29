import type {Op} from '../src/domain/edits';
import {docKey, journalFor, parseRecovery, recoveryFor, touchRecent} from '../src/domain/recovery';

const step = (text: string): Op[] => [{op: 'text', para: 0, start: 0, end: 0, text}];

describe('recovery', () => {
  it('keeps only the steps since the last save, plus typing in progress', () => {
    const r = recoveryFor('/d/a.docx', '10:1', [step('a'), step('b'), step('c')], 3, 1, step('typing'))!;
    expect(r.steps.map(s => (s[0] as {text: string}).text)).toEqual(['b', 'c', 'typing']);
  });

  it('has nothing to keep when saved, and cannot replay past an undone save', () => {
    expect(recoveryFor('/d/a.docx', '10:1', [step('a')], 1, 1, [])).toBeNull();
    expect(recoveryFor('/d/a.docx', '10:1', [step('a')], 1, -1, [])).toBeNull();
  });

  it('is offered only for the same file, unchanged', () => {
    const json = JSON.stringify(recoveryFor('/d/a.docx', '10:1', [step('a')], 1, 0, []));
    expect(parseRecovery(json, '/d/a.docx', '10:1')).not.toBeNull();
    expect(parseRecovery(json, '/d/a.docx', '11:2')).toBeNull();
    expect(parseRecovery('not json', '/d/a.docx', '10:1')).toBeNull();
  });
});

describe('recent documents and keys', () => {
  it('moves a document to the front and keeps ten', () => {
    let list = Array.from({length: 10}, (_, i) => ({path: `/p${i}`, name: `p${i}`, time: 0}));
    list = touchRecent(list, '/p5', 'p5');
    expect(list[0].path).toBe('/p5');
    expect(list).toHaveLength(10);
    expect(touchRecent(list, '/new', 'new')).toHaveLength(10);
  });
  it('makes stable, file-safe keys', () => {
    expect(docKey('/a/My Paper.docx')).toBe(docKey('/a/My Paper.docx'));
    expect(docKey('/a/My Paper.docx')).not.toBe(docKey('/b/My Paper.docx'));
    expect(docKey('/a/My Paper.docx')).toMatch(/^[A-Za-z0-9-]+$/);
  });
});

describe('the journal (audit 2026-09-28)', () => {
  const s = (t: string) => [{op: 'text', para: 0, start: 0, end: 0, text: t}] as Op[];

  it('is clean only when everything is saved', () => {
    expect(journalFor('/d/a.docx', '1', [s('a')], 1, 1, [], null)).toEqual({kind: 'clean'});
    expect(journalFor('/d/a.docx', '1', [s('a')], 1, 1, s('typing'), null)?.kind).toBe('record');
    // No fingerprint: leave whatever is recorded alone.
    expect(journalFor('/d/a.docx', null, [s('a')], 1, 0, [], null)).toBeNull();
  });

  it('after undoing past the save and editing on, records every step on the opened copy', () => {
    // Saved at 3; undone to 1; edited: a new step 2.
    const steps = [s('a'), s('b-new')];
    expect(journalFor('/d/a.docx', '1', steps, 2, 3, [], null)).toEqual({kind: 'needsBase', steps});
    const j = journalFor('/d/a.docx', '1', steps, 2, 3, s('more'), '/private/recovery/a.docx');
    expect(j).toMatchObject({kind: 'record', record: {base: '/private/recovery/a.docx', steps: [...steps, s('more')]}});
    // An overwritten save point (-1) is the same case.
    expect(journalFor('/d/a.docx', '1', steps, 2, -1, [], null)?.kind).toBe('needsBase');
  });
});
