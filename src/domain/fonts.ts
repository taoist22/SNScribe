// Stand-in fonts: free fonts with the same letter widths as common Word fonts, so a
// document's lines and pages break as in Word. The screen uses one when the document's own
// font is not on the Supernote but the stand-in is (installed in MyStyle/Fonts). Only the
// screen: the file keeps its fonts.

export const STAND_INS: Record<string, string> = {
  calibri: 'Carlito',
  cambria: 'Caladea',
  arial: 'Liberation Sans',
  helvetica: 'Liberation Sans',
  'times new roman': 'Liberation Serif',
  times: 'Liberation Serif',
  'courier new': 'Liberation Mono',
  courier: 'Liberation Mono',
  georgia: 'Gelasio',
};

export function standInFor(family: string): string | undefined {
  return STAND_INS[family.trim().toLowerCase()];
}

/** The families to load for a document: its own, and the stand-in of each. */
export function withStandIns(families: string[]): string[] {
  return [...new Set([...families, ...families.map(standInFor).filter((f): f is string => !!f)])];
}

/** The family the screen draws `family` in: itself if loaded, else its loaded stand-in, else none (the default font). */
export function shownFont(family: string | undefined, loaded: Set<string> | undefined): string | undefined {
  if (!family || !loaded) {
    return undefined;
  }
  if (loaded.has(family)) {
    return family;
  }
  const s = standInFor(family);
  return s && loaded.has(s) ? s : undefined;
}
