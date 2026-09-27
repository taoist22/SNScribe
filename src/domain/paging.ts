// Page breaks for the reader. Pure: works on measured boxes, no React.
//
// The reader draws a "window" of blocks starting at the page's anchor block as one column,
// shifted up by the anchor's offset. A page shows column y in [pageTop, pageTop + height).
// The next page starts at the top of the first LINE that does not fit — never mid-line —
// so a paragraph longer than a page flows across pages.

/** A laid-out line: top and height (relative to its block), and how many characters it holds. */
export type LineBox = {y: number; height: number; len?: number};
/** A block's frame in column coordinates; `lines` (relative to the block) for text blocks. */
export type BlockBox = {top: number; height: number; lines?: LineBox[]; forced?: boolean};

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
    // A page break before this block: the page ends here, whatever room is left.
    if (box.forced && box.top > pageTop) {
      return {kind: 'at', index: i, top: box.top};
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

/** Where a page begins: its anchor, and the character of the anchor's paragraph it starts at. */
export type PageStart = {anchor: Anchor; char: number};

/**
 * Every page break inside one measured window of blocks that starts at `start` (window
 * block 0 = start.block). Uses the same findBreak as turning pages, so the map matches
 * what ▶ shows when paging from the start.
 *
 * Returns the pages that begin inside the window, then either `done` (the document ends in
 * this window) or where the next window must begin: the start of the page that runs past
 * the window's end, so that page is measured whole next time. `null` result = a block is
 * not measured yet.
 */
export function breaksInWindow(
  boxes: Array<BlockBox | undefined>,
  start: Anchor,
  pageHeight: number,
  lastWindow: boolean,
): {pages: PageStart[]; next: Anchor | null} | null {
  const pages: PageStart[] = [];
  let top = start.offset;
  let current = start;
  for (let guard = 0; guard < 10000; guard++) {
    const br = findBreak(boxes, top, pageHeight);
    if (br.kind === 'pending') {
      return null;
    }
    if (br.kind === 'after') {
      return {pages, next: lastWindow ? null : current};
    }
    const box = boxes[br.index]!;
    const anchor = {block: start.block + br.index, offset: Math.max(0, br.top - box.top)};
    let char = 0;
    for (const line of box.lines ?? []) {
      if (box.top + line.y >= br.top) {
        break;
      }
      char += line.len ?? 0;
    }
    pages.push({anchor, char});
    current = anchor;
    top = br.top;
  }
  return {pages, next: lastWindow ? null : current};
}

/** Page index (0-based) of `at` in a page map: the last page that begins at or before it. */
export function pageIndexOf(pages: PageStart[], at: Anchor): number {
  let lo = 0;
  let hi = pages.length - 1;
  let found = 0;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const a = pages[mid].anchor;
    if (a.block < at.block || (a.block === at.block && a.offset <= at.offset + 1)) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}
