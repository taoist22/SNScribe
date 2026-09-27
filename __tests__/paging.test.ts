import {anchorAfter, findBreak, windowEnd, type BlockBox} from '../src/domain/paging';

const lines = (n: number, h = 30): Array<{y: number; height: number}> =>
  Array.from({length: n}, (_, i) => ({y: i * h, height: h}));

describe('findBreak', () => {
  it('breaks before the first block that does not fit', () => {
    const boxes: BlockBox[] = [
      {top: 0, height: 100},
      {top: 120, height: 100},
    ];
    expect(findBreak(boxes, 0, 200)).toEqual({kind: 'at', index: 1, top: 120});
  });

  it('breaks at the first line that does not fit inside a long paragraph', () => {
    const boxes: BlockBox[] = [{top: 0, height: 300, lines: lines(10)}];
    // Page 0–100: lines 0,1,2 fit (0–90); line 3 (90–120) does not.
    expect(findBreak(boxes, 0, 100)).toEqual({kind: 'at', index: 0, top: 90});
    // Next page from 90: lines 3,4,5 (90–180) fit in 90–190; line 6 (180–210) does not.
    expect(findBreak(boxes, 90, 100)).toEqual({kind: 'at', index: 0, top: 180});
  });

  it('cuts a block with no lines that is taller than the page', () => {
    expect(findBreak([{top: 0, height: 500}], 0, 200)).toEqual({kind: 'at', index: 0, top: 200});
  });

  it('reports pending until every block is measured, and after when all fit', () => {
    expect(findBreak([{top: 0, height: 10}, undefined], 0, 200)).toEqual({kind: 'pending'});
    expect(findBreak([{top: 0, height: 10}, {top: 20, height: 10}], 0, 200)).toEqual({kind: 'after', top: 30});
  });
});

describe('anchorAfter', () => {
  const boxes: BlockBox[] = [
    {top: 0, height: 100},
    {top: 120, height: 300, lines: lines(10)},
  ];
  it('turns a break into an anchor inside the block', () => {
    expect(anchorAfter({block: 7, offset: 0}, boxes, {kind: 'at', index: 1, top: 180})).toEqual({block: 8, offset: 60});
  });
  it('moves past the window when everything fit', () => {
    expect(anchorAfter({block: 7, offset: 0}, boxes, {kind: 'after', top: 420})).toEqual({block: 9, offset: 0});
    expect(anchorAfter({block: 7, offset: 0}, boxes, {kind: 'pending'})).toBeNull();
  });
});

describe('windowEnd', () => {
  it('takes at least one block and stops after the word budget', () => {
    expect(windowEnd([5000, 10], 0)).toBe(1);
    expect(windowEnd([400, 400, 400, 400], 0)).toBe(3);
    expect(windowEnd([1, 1, 1], 1, 100, 1)).toBe(2);
  });
});

import {breaksInWindow, pageIndexOf} from '../src/domain/paging';

describe('breaksInWindow', () => {
  const lines = (n: number, h = 30, len = 40) => Array.from({length: n}, (_, i) => ({y: i * h, height: h, len}));

  it('finds every page start in a window, with the character each begins at', () => {
    // 10 lines of 30 = 300 tall; pages of 100 hold 3 lines each.
    const boxes = [{top: 0, height: 300, lines: lines(10)}];
    const r = breaksInWindow(boxes, {block: 4, offset: 0}, 100, true)!;
    expect(r.pages.map(p => p.anchor)).toEqual([
      {block: 4, offset: 90},
      {block: 4, offset: 180},
      {block: 4, offset: 270},
    ]);
    expect(r.pages.map(p => p.char)).toEqual([120, 240, 360]);
    expect(r.next).toBeNull();
  });

  it('asks for a new window at the page that runs past the end', () => {
    const boxes = [{top: 0, height: 60}, {top: 70, height: 60}];
    const r = breaksInWindow(boxes, {block: 0, offset: 0}, 100, false)!;
    expect(r.pages).toEqual([{anchor: {block: 1, offset: 0}, char: 0}]);
    expect(r.next).toEqual({block: 1, offset: 0});
  });

  it('waits until everything is measured', () => {
    expect(breaksInWindow([undefined], {block: 0, offset: 0}, 100, true)).toBeNull();
  });
});

describe('pageIndexOf', () => {
  const pages = [0, 5, 9].map(block => ({anchor: {block, offset: 0}, char: 0}));
  it('finds the page a position is on', () => {
    expect(pageIndexOf(pages, {block: 0, offset: 0})).toBe(0);
    expect(pageIndexOf(pages, {block: 7, offset: 40})).toBe(1);
    expect(pageIndexOf(pages, {block: 9, offset: 0})).toBe(2);
    expect(pageIndexOf(pages, {block: 20, offset: 0})).toBe(2);
  });
});
