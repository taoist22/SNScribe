// Paper formats: one tap turns a whole document into APA 7, MLA 9 or Chicago (Turabian)
// form, as ordinary edits (one undo step): fonts, spacing, margins, headings, reference
// lists, and the page number top right (after the last name, in MLA).

import type {Op} from './edits';
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
      const centered = p.kind === 'title' || isRefHeading || (p.kind === 'heading' && p.level <= 1 && format !== 'mla');
      ops.push({op: 'para', para: p.index, align: centered ? 'center' : 'left', line: DOUBLE, lineRule: 'auto', before: 0, after: 0, first: 0});
      // APA and Chicago headings are bold; MLA titles are plain.
      if (len > 0 && format !== 'mla') {
        ops.push({op: 'format', para: p.index, start: 0, end: len, prop: 'b', on: true});
      }
      if (len > 0 && format === 'mla') {
        ops.push({op: 'format', para: p.index, start: 0, end: len, prop: 'b', on: false});
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
    // Body text; list items keep their own indentation.
    ops.push({op: 'para', para: p.index, line: DOUBLE, lineRule: 'auto', before: 0, after: 0, ...(p.num ? {} : {first: 720})});
  }
  return ops;
}
