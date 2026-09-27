// The document model the native reader (DocxModule.open → DocxReader) produces.
//
// Paragraph text is the concatenation of its runs' `t`. Inline objects (images, note
// references, embedded objects) are one U+FFFC each, so character offsets match the XML
// the writer edits.

export const OBJECT = '￼';

export type Run = {
  t: string;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  /** Highlighted in Word. */
  h?: boolean;
  /** Inside a hyperlink. */
  l?: boolean;
  sup?: boolean;
  obj?: 'image' | 'note' | 'object';
  /** Inside a field or content control: formattable, but its text is not editable. */
  k?: boolean;
  /** Font family (theme fonts resolved). */
  f?: string;
  /** Size in half-points (w:sz): 22 = 11 pt. */
  sz?: number;
  /** Inside a tracked insertion: its id (ParagraphBlock.revs). */
  rv?: string;
};

/**
 * A tracked change. 'ins': the runs whose rv is this id. 'del': deleted text that is not
 * part of the paragraph's text; it shows (struck through) before character `at`.
 */
export type Revision = {
  id: string;
  kind: 'ins' | 'del';
  author: string;
  date: string;
  at?: number;
  runs?: Run[];
  move?: boolean;
};

/** A comment anchor at text offset `at`: where its range starts, ends, and its reference mark. */
export type Mark = {id: string; kind: 'start' | 'end' | 'ref'; at: number};

/** A Word comment. `parent`: the comment it replies to. `pictures`: pictures in it (a handwritten note). */
export type Comment = {
  id: string;
  author: string;
  initials: string;
  date: string;
  text: string;
  parent?: string;
  done?: boolean;
  pictures?: number;
};

export type ParagraphBlock = {
  type: 'p';
  /** Comment anchors in its text. */
  marks?: Mark[];
  /** Tracked insertions and deletions in its text, in reading order. */
  revs?: Revision[];
  /** Ordinal among the body's top-level paragraphs: the address edits use. */
  index: number;
  style: string;
  kind: 'title' | 'subtitle' | 'heading' | 'body';
  level: number;
  align: 'left' | 'center' | 'right' | 'justify';
  /** Left indent in twips (1/1440 inch). */
  indent: number;
  /** List label ("1.", "•"), when the paragraph is in a list. */
  list?: string;
  runs: Run[];
  /** Its paragraph mark ends a section: nothing may be joined onto it. */
  sect?: boolean;
  /** The list it belongs to — a document list id (number) or a list made while editing (string) — and its level. */
  num?: {id: number | string; lvl: number};
  /** Space before / after (twips) and line spacing: 240ths of a line when lineRule is 'auto', else twips. */
  before?: number;
  after?: number;
  line?: number;
  lineRule?: string;
  /** First-line indent in twips; negative = hanging. */
  first?: number;
  /** Starts on a new page. */
  pb?: boolean;
  /** Font and size (half-points) text here has unless it sets its own: from the style and document defaults. */
  bf?: string;
  bs?: number;
};

export type TableBlock = {type: 'table'; rows: number; cols: number; preview: string};
export type ProtectedBlock = {type: 'protected'; what: string; preview: string};
export type Block = ParagraphBlock | TableBlock | ProtectedBlock;

export type Report = {
  paragraphs: number;
  tables: number;
  images: number;
  trackedChanges: number;
  comments: number;
  fields: number;
  contentControls: number;
};

export type PageSetup = {width: number; height: number; top: number; right: number; bottom: number; left: number; landscape: boolean};

export type ListLevel = {fmt: string; text: string; start: number} | null;
/** A list's levels, and the numbers it restarts at. */
export type ListDef = {levels: ListLevel[]; starts: Record<string, number>};

export type DocxDocument = {
  path: string;
  name: string;
  bytes: number;
  ms: number;
  report: Report;
  blocks: Block[];
  /** Definitions of the lists the paragraphs use, by id. */
  lists: Record<string, ListDef>;
  /** Page size and margins (twips) of the last section. */
  page?: PageSetup;
  /** Tracked formatting and paragraph changes DOCX can't review; they stay as they are. */
  otherRevisions?: number;
  comments?: Comment[];
  /**
   * The default header and footer of the last section, as DOCX can edit them. `other`: it
   * also has a logo or table, which stays; text and alignment are then of its page-number line.
   */
  header?: {text: string; pageNumber: boolean; align: 'left' | 'center' | 'right' | 'justify'; other?: boolean};
  footer?: {text: string; pageNumber: boolean; align: 'left' | 'center' | 'right' | 'justify'; other?: boolean};
  /**
   * Set for a document made with New: edits are applied to `source` (a pristine blank in
   * private storage) and saved over `saveTo` (the new file), instead of an "-edited" copy.
   */
  source?: string;
  saveTo?: string;
};

/** Every font family the document's text uses. */
export function fontsUsed(blocks: Block[]): string[] {
  const out = new Set<string>();
  for (const b of blocks) {
    if (b.type === 'p') {
      if (b.bf) {
        out.add(b.bf);
      }
      for (const r of b.runs) {
        if (r.f) {
          out.add(r.f);
        }
      }
    }
  }
  return [...out].sort();
}

export function paragraphText(p: ParagraphBlock): string {
  return p.runs.map(r => r.t).join('');
}

export function wordCount(b: Block): number {
  if (b.type !== 'p') {
    return 20;
  }
  const text = paragraphText(b).trim();
  return text ? text.split(/\s+/).length : 1;
}

/** Headings and titles, for the contents list. */
export function outline(blocks: Block[]): Array<{block: number; level: number; text: string}> {
  const out: Array<{block: number; level: number; text: string}> = [];
  blocks.forEach((b, i) => {
    if (b.type === 'p' && b.kind !== 'body') {
      const text = paragraphText(b).replace(/\s+/g, ' ').trim();
      if (text) {
        out.push({block: i, level: b.kind === 'heading' ? b.level : b.kind === 'title' ? 0 : 1, text});
      }
    }
  });
  return out;
}

/** Words and characters (objects not counted), for Word count. */
export function countWords(texts: string[]): {words: number; chars: number} {
  let words = 0;
  let chars = 0;
  for (const t of texts) {
    const clean = t.replace(/\ufffc/g, '');
    chars += clean.replace(/\s/g, '').length;
    words += clean.split(/\s+/).filter(w => /[\p{L}\p{N}]/u.test(w)).length;
  }
  return {words, chars};
}
