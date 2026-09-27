// List numbers, counted the way the native reader counts them (DocxReader.Numbering), so
// the screen can renumber after edits: in document order, per list, a level's counter
// restarting whenever a higher level counts on.

import type {Block, ListDef, ParagraphBlock} from '../model/docx';

const BULLETS = ['•', '◦', '▪'];

/** The definition DocxEditor.Lists gives a list made while editing ("n…" numbers, "b…" bullets). */
export function newListDef(kind: 'number' | 'bullet'): ListDef {
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
  return {
    levels: Array.from({length: 9}, (_, i) =>
      kind === 'bullet' ? {fmt: 'bullet', text: BULLETS[i % 3], start: 1} : {fmt: fmts[i % 3], text: `%${i + 1}.`, start: 1},
    ),
    starts: {0: 1},
  };
}

export function listKind(id: number | string): 'number' | 'bullet' | null {
  return typeof id === 'string' ? (id.startsWith('b') ? 'bullet' : 'number') : null;
}

function definition(lists: Record<string, ListDef>, id: number | string): ListDef | null {
  const kind = listKind(id);
  return lists[String(id)] ?? (kind ? newListDef(kind) : null);
}

function letters(n: number): string {
  if (n <= 0) {
    return String(n);
  }
  return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n: number): string {
  if (n <= 0 || n >= 4000) {
    return String(n);
  }
  const table: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let rest = n;
  let out = '';
  for (const [v, s] of table) {
    while (rest >= v) {
      out += s;
      rest -= v;
    }
  }
  return out;
}

export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'lowerLetter':
      return letters(n).toLowerCase();
    case 'upperLetter':
      return letters(n);
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'upperRoman':
      return roman(n);
    case 'decimalZero':
      return String(n).padStart(2, '0');
    default:
      return String(n);
  }
}

/** Every list paragraph's label, recounted. Paragraphs whose label is unchanged keep their identity. */
export function recount(blocks: Block[], lists: Record<string, ListDef>): Block[] {
  const counters = new Map<string, number[]>();
  return blocks.map(b => {
    if (b.type !== 'p' || !b.num) {
      return b;
    }
    const def = definition(lists, b.num.id);
    const i = Math.max(0, Math.min(8, b.num.lvl));
    const level = def?.levels[i];
    if (!def || !level) {
      return b.list === undefined ? b : {...b, list: undefined};
    }
    const key = String(b.num.id);
    const count = counters.get(key) ?? Array(9).fill(-1);
    counters.set(key, count);
    count[i] = count[i] < 0 ? def.starts[i] ?? level.start : count[i] + 1;
    for (let deeper = i + 1; deeper < 9; deeper++) {
      count[deeper] = -1;
    }
    let label: string;
    if (level.fmt === 'none') {
      label = '';
    } else if (level.fmt === 'bullet') {
      label = BULLETS[i % BULLETS.length];
    } else {
      label = level.text.replace(/%([1-9])/g, (_, d: string) => {
        const k = Number(d) - 1;
        const l = def.levels[k];
        const n = count[k] < 0 ? def.starts[k] ?? l?.start ?? 1 : count[k];
        return formatNumber(n, l?.fmt ?? 'decimal');
      });
    }
    return label === b.list ? b : ({...b, list: label} as ParagraphBlock);
  });
}
