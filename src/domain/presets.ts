// Paper formats: one tap turns a whole document into APA 7, MLA 9 or Chicago (Turabian)
// form, as ordinary edits (one undo step): fonts, spacing, margins, headings, reference
// lists, and the page number top right (after the last name, in MLA).

import type {Op, StyleKind} from './edits';
import {paragraphText, type Block, type ParagraphBlock} from '../model/docx';

export type PaperFormat = 'apa' | 'mla' | 'chicago';

export const PAPER_FORMATS: Array<{id: PaperFormat; name: string; summary: string}> = [
  {id: 'apa', name: 'APA 7', summary: 'Times New Roman 12, double spaced, 0.5″ first-line indents, bold headings (level 1 centered), 1″ margins, page numbers top right, hanging indents under References.'},
  {id: 'mla', name: 'MLA 9', summary: 'Times New Roman 12, double spaced, 0.5″ first-line indents, centered title, 1″ margins, last name and page number top right, hanging indents under Works Cited.'},
  {id: 'chicago', name: 'Chicago', summary: 'Times New Roman 12, double-spaced text, 0.5″ first-line indents, centered level-1 headings, 1″ margins, page numbers top right; Bibliography single spaced with a blank line between entries, hanging.'},
];

const FONT = 'Times New Roman';
const SIZE = 24; // 12 pt
const INCH = 1440;
const DOUBLE = 480;

type Level = {align: 'left' | 'center'; bold: boolean; italic: boolean};

/**
 * Heading levels 1–3 in each format, all at body size. APA 7: centred bold / left bold /
 * left bold italic. MLA 9 (its handbook's suggested scheme): left bold / left italic /
 * centred bold. Chicago (Turabian): centred bold / centred plain / left bold.
 */
export const HEADING_LEVELS: Record<PaperFormat, [Level, Level, Level]> = {
  apa: [
    {align: 'center', bold: true, italic: false},
    {align: 'left', bold: true, italic: false},
    {align: 'left', bold: true, italic: true},
  ],
  mla: [
    {align: 'left', bold: true, italic: false},
    {align: 'left', bold: false, italic: true},
    {align: 'center', bold: true, italic: false},
  ],
  chicago: [
    {align: 'center', bold: true, italic: false},
    {align: 'center', bold: false, italic: false},
    {align: 'left', bold: true, italic: false},
  ],
};

/** The heading that starts a reference list in each format. */
const REFERENCES: Record<PaperFormat, RegExp> = {
  apa: /^references$/i,
  mla: /^works cited$/i,
  chicago: /^(bibliography|references)$/i,
};

export function presetOps(blocks: Block[], format: PaperFormat, lastName = ''): Op[] {
  const ops: Op[] = [
    {op: 'page', para: -1, top: INCH, right: INCH, bottom: INCH, left: INCH},
    {op: 'defaults', para: -1, font: FONT, size: SIZE},
    // Page numbers top right on every page; MLA puts the writer's last name before it.
    {op: 'headerFooter', para: -1, kind: 'header', text: format === 'mla' ? lastName.trim() : '', pageNumber: true, align: 'right'},
    // The heading styles themselves, so headings added later come out right too.
    {
      op: 'styleDefs',
      para: -1,
      defs: HEADING_LEVELS[format].map((l, i) => ({kind: `heading${i + 1}` as StyleKind, font: FONT, size: SIZE, bold: l.bold, italic: l.italic, align: l.align, line: DOUBLE})),
    },
  ];
  let inReferences = false;
  for (const b of blocks) {
    if (b.type !== 'p') {
      continue;
    }
    const p: ParagraphBlock = b;
    const text = paragraphText(p);
    const len = text.length;
    const isRefHeading = REFERENCES[format].test(text.trim());
    const heading = p.kind !== 'body' || isRefHeading;
    if (len > 0) {
      ops.push({op: 'runStyle', para: p.index, start: 0, end: len, font: FONT, size: SIZE});
    }
    if (heading) {
      inReferences = isRefHeading;
      // Titles and the reference-list heading are centred (bold except in MLA); headings
      // follow their level in the format.
      const level = p.kind === 'heading' && !isRefHeading ? HEADING_LEVELS[format][Math.min(3, Math.max(1, p.level)) - 1] : null;
      const align = level ? level.align : 'center';
      const bold = level ? level.bold : format !== 'mla';
      const italic = level ? level.italic : false;
      ops.push({op: 'para', para: p.index, align, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0});
      if (len > 0) {
        ops.push({op: 'format', para: p.index, start: 0, end: len, prop: 'b', on: bold});
        ops.push({op: 'format', para: p.index, start: 0, end: len, prop: 'i', on: italic});
      }
      continue;
    }
    if (inReferences) {
      // Reference entries: hanging 0.5″; Chicago's bibliography is single spaced with a blank line between.
      ops.push(
        format === 'chicago'
          ? {op: 'para', para: p.index, line: 240, lineRule: 'auto', before: 0, after: 240, first: -720}
          : {op: 'para', para: p.index, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: -720},
      );
      continue;
    }
    // A block quotation: its whole-block indent stays, with no first-line indent.
    if (p.quote) {
      ops.push({op: 'para', para: p.index, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0});
      continue;
    }
    // Body text; list items keep their own indentation.
    ops.push({op: 'para', para: p.index, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, ...(p.num ? {} : {first: 720})});
  }
  return ops;
}
