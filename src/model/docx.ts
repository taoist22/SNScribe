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
};

export type ParagraphBlock = {
  type: 'p';
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
};

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
