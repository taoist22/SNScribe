// Page breaks for the reader. Pure: works on measured boxes, no React.
//
// The reader draws a "window" of blocks starting at the page's anchor block as one column,
// shifted up by the anchor's offset. A page shows column y in [pageTop, pageTop + height).
// The next page starts at the top of the first LINE that does not fit — never mid-line —
// so a paragraph longer than a page flows across pages.

export type LineBox = {y: number; height: number};
/** A block's frame in column coordinates; `lines` (relative to the block) for text blocks. */
export type BlockBox = {top: number; height: number; lines?: LineBox[]};

/** Page top: a block index into the document and a y offset inside that block. */
export type Anchor = {block: number; offset: number};

export type Break =
  | {kind: 'at'; index: number; top: number} // next page starts in window block `index`, column y `top`
  | {kind: 'after'; top: number} // everything rendered fits; next page starts after the window
  | {kind: 'pending'}; // a block is not measured yet

/**
 * Where the next page starts. `boxes[i]` is window block i (window block 0 is the anchor
 * block, top 0). A block taller than the page with no line boxes is cut at the page bottom.
 */
export function findBreak(boxes: Array<BlockBox | undefined>, pageTop: number, pageHeight: number): Break {
  const bottom = pageTop + pageHeight;
  let lastBottom = pageTop;
  for (let i = 0; i < boxes.length; i++) {
    const box = boxes[i];
    if (!box) {
      return {kind: 'pending'};
    }
    if (box.top + box.height <= bottom) {
      lastBottom = Math.max(lastBottom, box.top + box.height);
      continue;
    }
    if (box.lines && box.lines.length > 0) {
      for (const line of box.lines) {
        const lineTop = box.top + line.y;
        if (lineTop + line.height > bottom) {
          return {kind: 'at', index: i, top: lineTop > pageTop ? lineTop : bottom};
        }
      }
      // All lines fit; only trailing space overflows. Start at the next block.
      lastBottom = bottom;
      continue;
    }
    return {kind: 'at', index: i, top: box.top > pageTop ? box.top : bottom};
  }
  return {kind: 'after', top: lastBottom};
}

/** The anchor for the page after a break, given the current anchor and window. */
export function anchorAfter(anchor: Anchor, boxes: Array<BlockBox | undefined>, br: Break): Anchor | null {
  if (br.kind === 'pending') {
    return null;
  }
  if (br.kind === 'after') {
    return {block: anchor.block + boxes.length, offset: 0};
  }
  const box = boxes[br.index]!;
  return {block: anchor.block + br.index, offset: Math.max(0, br.top - box.top)};
}

/** The window of blocks drawn from `start`: enough words for more than one page. */
export function windowEnd(wordCounts: number[], start: number, budget = 1100, maxBlocks = 80): number {
  let words = 0;
  let end = start;
  while (end < wordCounts.length && (end === start || (words < budget && end - start < maxBlocks))) {
    words += wordCounts[end];
    end++;
  }
  return end;
}
