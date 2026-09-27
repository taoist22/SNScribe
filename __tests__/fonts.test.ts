import {shownFont, withStandIns} from '../src/domain/fonts';

describe('stand-in fonts', () => {
  it('loads each font and its stand-in', () => {
    expect(withStandIns(['Calibri', 'JetBrains Mono'])).toEqual(['Calibri', 'JetBrains Mono', 'Carlito']);
  });

  it('shows the font itself, else its loaded stand-in, else nothing', () => {
    const loaded = new Set(['Carlito', 'Times New Roman']);
    expect(shownFont('Calibri', loaded)).toBe('Carlito');
    expect(shownFont('Times New Roman', loaded)).toBe('Times New Roman');
    expect(shownFont('Cambria', loaded)).toBeUndefined();
    expect(shownFont('Aptos', loaded)).toBeUndefined();
    expect(shownFont(undefined, loaded)).toBeUndefined();
  });
});
