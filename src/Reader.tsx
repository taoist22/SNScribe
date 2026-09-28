import React, {useCallback, useEffect, useMemo, useRef, useState} from 'react';
import {
  DeviceEventEmitter,
  Image,
  Keyboard,
  PixelRatio,
  Pressable,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  View,
  findNodeHandle,
  type GestureResponderEvent,
  type LayoutChangeEvent,
} from 'react-native';
import {PluginManager, RattaFileSelector} from 'sn-plugin-lib';
import {BlockView} from './BlockView';
import {PageCounter} from './PageCounter';
import {pageInfo, pageMarks} from './domain/pageInfo';
import {targetLabel, wordCounts} from './domain/wordcount';
import {figureOps, pictureEmu} from './domain/figures';
import {TouchLayer} from './services/touch';
import {
  applyOps,
  comparePos,
  deletionRange,
  expectedTexts,
  commentsAfter,
  findMatches,
  formatOps,
  fromShown,
  HIGHLIGHTS,
  highlightOps,
  headerFooterAfter,
  inksAfter,
  joinProblem,
  linkProblem,
  linkSpan,
  linkUrl,
  looksAfter,
  nextCommentId,
  penSelection,
  pageAfter,
  replaceAllOps,
  rangesBetween,
  splitProblem,
  styleOps,
  textEditProblem,
  threadIds,
  toShown,
  pageBreakChars,
  wordAround,
  type FormatProp,
  type Op,
  type Pos,
  type Range,
  type StyleKind,
} from './domain/edits';
import {listKind, recount} from './domain/lists';
import {PAPER_FORMATS, presetOps} from './domain/presets';
import {docKey, parseRecovery, recoveryFor, touchRecent, type RecentDoc, type RecoveryRecord} from './domain/recovery';
import {DocxInk, InkSurfaceView, activateInk, deactivateInk, isInkAvailable} from './services/ink';
import {shownFont, withStandIns} from './domain/fonts';
import {CITE_STYLES, htmlToPieces, inText, piecesText, referenceOps, type CiteStyle, type Source} from './domain/citations';
import {credentialsIn, searchZotero, setZoteroLog, testZotero, zoteroItem, type ZoteroAccount} from './services/zotero';
import {lookUpDoi} from './services/doi';
import {QUOTES_KEY, SOURCES_KEY, fileName, isEpub, parseList, parseMap, quoteWithReference, type SavedQuote, type SourceRef} from './domain/quotes';
import {detailsFromEpub, formatSource, type Details} from './domain/reference';
import {misspelledRanges, normal, tokens} from './domain/spelling';
import {anchorAfter, findBreak, pageIndexOf, pageOfChar, windowEnd, type Anchor, type BlockBox, type Break, type LineBox, type PageStart} from './domain/paging';
import {countWords, fontsUsed, outline, paragraphText, wordCount, type Block, type DocxDocument, type ParagraphBlock, type Run} from './model/docx';
import {ensureFileReadPermission, ensureFileWritePermission, ensureInternetPermission} from './pluginPermissions';
import {Docx, DocxKeys, DocxText, errorText, log, nativeBuild, type KeyPress} from './services/native';

/**
 * Read a Word document a page at a time (M1) and mark it up (M2).
 *
 * Paging: the blocks from the page's anchor are drawn as one column, shifted up by the
 * anchor's offset, inside a clipped page. Once every block reports its frame (and every
 * paragraph its lines), findBreak picks where the next page starts — at a line boundary —
 * and a white mask hides the part-line below it. ▶ moves the anchor there; ◀ goes back
 * through the pages already seen.
 *
 * Editing: the pen selects words (the paragraph's own TextView reports the character under
 * the pen, via DocxText); a tap places a caret, a double tap selects a word. One invisible
 * text field keeps the keyboard (Bluetooth or on-screen, incl. handwriting) connected
 * from the first tap until the keyboard is hidden (Edit ▸ Hide keyboard), and the first
 * key decides what it does: typing at the caret, typing over or deleting the selection,
 * Enter for a new paragraph. Typed text shows in the page as it is typed — a pending edit on top of the committed ones — and becomes
 * one undo step when typing ends. Buttons turn
 * the selection into edits (domain/edits), one undo step per button press. The page shows
 * the original with the first `cursor` steps applied — Undo/Redo move the cursor — and Save
 * hands the same edits to the native writer, which saves a verified copy beside the
 * original. The original file is never changed.
 */

const HEADER_H = 64;
const BAR_H = 60;
const PAD = 16;
/** The margin column beside the text: tracked changes and comments, as in Word's margin. */
const MARGIN_W = 290;
const MARGIN_GAP = 14;
/** A margin card's height: cards are a fixed size so they can be stacked before they are drawn. */
const CARD_H = 84;
/** The pen pad's header row (fixed: the surface below it must never move). */
const PAD_HEAD = 64;
/** The pen pad's writing column: margin-shaped. */
const PAD_W = 420;
/** The margin column folded away: a strip with a count and a button to open it again. */
const MARGIN_STRIP = 44;
/** Ink colours for the note picture in Word (the pad itself can only show black). */
const INK_COLORS: Array<[string, string]> = [
  ['Black', '#000000'],
  ['Blue', '#1F4FD8'],
  ['Red', '#C62828'],
  ['Green', '#2E7D32'],
  ['Purple', '#6A1B9A'],
];
const START: Anchor = {block: 0, offset: 0};
/** Pen travel below this (dp) is a tap: it selects the word under the pen. */
const TAP_SLOP = 12;
/** A finger swipe this far (dp), mostly sideways, turns the page. */
const SWIPE_MIN = 80;
/** A second tap this soon, this close, selects the word under it. */
const DOUBLE_TAP_MS = 500;

type Frame = {x: number; top: number; height: number};
type Selection = {from: Pos; to: Pos};
type Typing = {mode: 'insert'; at: Pos} | {mode: 'replace'; range: Range};
type HfLine = {text: string; align: 'left' | 'center' | 'right'; page: boolean};

type Menu =
  | 'file' | 'edit' | 'style' | 'list' | 'font' | 'size' | 'name' | 'folder' | 'view' | 'para' | 'count' | 'page' | 'preset' | 'header' | 'link' | 'thread' | 'comment' | 'note' | 'hl' | 'find' | 'cite' | 'quotes' | 'spell' | 'goto' | 'picture'
  | 'recent' | 'versions' | 'recover';

/** Drop-down menu width; menus are kept inside the screen. */
const MENU_W = 340;

const STORAGE = '/storage/emulated/0';
const DOCUMENTS = `${STORAGE}/Document`;
/** Reader text sizes (A− / A+), as factors of the document's own sizes. */
const SCALES = [0.8, 0.9, 1, 1.15, 1.3, 1.5, 1.75];

/** One line of text: line breaks and other control characters become spaces. */
const clean = (text: string) => text.replace(/[\r\n\u0000-\u0008\u000b-\u001f\ufffc]/g, ' ');

/** The first block index where two versions of the text differ (edits replace only the blocks they change). */
function firstDifference(a: Block[], b: Block[]): number {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) {
    i++;
  }
  return i;
}

const typedRange = (t: Typing): Range =>
  t.mode === 'insert' ? {para: t.at.para, start: t.at.offset, end: t.at.offset} : t.range;

export function Reader(): React.JSX.Element {
  const [doc, setDoc] = useState<DocxDocument | null>(null);
  const [anchor, setAnchor] = useState<Anchor>(START);
  const [history, setHistory] = useState<Anchor[]>([]);
  /** Pages view: only the pages with comments, handwritten notes or tracked changes. */
  const [notesOnly, setNotesOnly] = useState(false);
  const [goToText, setGoToText] = useState('');
  /** Insert picture…: the file chosen, its pixels, and the figure label/title to add. */
  const [pic, setPic] = useState<{para: number; path?: string; w?: number; h?: number; label: boolean; title: string; note: string}>({
    para: -1,
    label: true,
    title: '',
    note: '',
  });
  /** Word-count target for this document (body text), kept per document; null = none. */
  const [target, setTarget] = useState<number | null>(null);
  const [targetText, setTargetText] = useState('');
  const [goToNote, setGoToNote] = useState('');
  const [status, setStatus] = useState('');
  const [busy, setBusy] = useState(false);
  const [contents, setContents] = useState(false);
  /** The open drop-down menu, and the x of the button that opened it. */
  const [menu, setMenu] = useState<Menu | null>(null);
  const [menuX, setMenuX] = useState(0);
  /** Font families loaded on this device (from MyStyle/Fonts, or added by hand). */
  const [fonts, setFonts] = useState<Set<string>>(new Set());
  /** Naming a new document: the name being typed, or null. */
  const [naming, setNaming] = useState<string | null>(null);
  /** New from template: the .docx/.dotx the new document is copied from. */
  const [template, setTemplate] = useState<string | null>(null);
  /** The writer's last name, for MLA's header (remembered). */
  const [lastName, setLastName] = useState('');
  /** Header & footer form: each line as edited, and as it was when the form opened. */
  const [hfForm, setHfForm] = useState<{header: HfLine; footer: HfLine; was: {header: HfLine; footer: HfLine}} | null>(null);
  const [linkForm, setLinkForm] = useState<{para: number; start: number; end: number; url: string} | null>(null);
  /** The comment thread open in the panel (its first comment's id). */
  const [thread, setThread] = useState<string | null>(null);
  const [replyText, setReplyText] = useState('');
  const [commentForm, setCommentForm] = useState<{fromPara: number; from: number; toPara: number; to: number; text: string; quote: string} | null>(null);
  /** The pen pad, open for a note on these words. */
  const [pad, setPad] = useState<{para: number; at: number; quote: string} | null>(null);
  const [noteOpen, setNoteOpen] = useState<string | null>(null);
  const [inkOk, setInkOk] = useState(false);
  /** The name comments are signed with (remembered). */
  const [author, setAuthor] = useState('');
  /** Zotero: the account (stored privately), the citation style, and the Cite panel's state. */
  const [zotero, setZotero] = useState<ZoteroAccount | null>(null);
  const [citeStyle, setCiteStyle] = useState<CiteStyle>('apa');
  const [cite, setCite] = useState<{
    setup: boolean;
    userId: string;
    apiKey: string;
    query: string;
    results: Source[] | null;
    chosen: Source | null;
    narrative: boolean;
    page: string;
    busy: boolean;
    /** The last problem or result, shown in the panel itself. */
    message: string;
  }>({setup: false, userId: '', apiKey: '', query: '', results: null, chosen: null, narrative: false, page: '', busy: false, message: ''});
  useEffect(() => setZoteroLog(line => log(line)), []);

  // What Android reports about keyboards, for the strip that sometimes covers the page's
  // last lines after switching from the on-screen keyboard to the Bluetooth one (CT).
  useEffect(() => {
    const shown = Keyboard.addListener('keyboardDidShow', e =>
      log(`keyboard shown: ${Math.round(e.endCoordinates.height)} dp high, top at ${Math.round(e.endCoordinates.screenY)}`),
    );
    const hidden = Keyboard.addListener('keyboardDidHide', () => log('keyboard hidden'));
    return () => {
      shown.remove();
      hidden.remove();
    };
  }, []);
  /** Spelling: marks while typing (View), the language, the dictionary's answers so far, the user's words. */
  const SPELL_LANG = 'en_US';
  const [spellOn, setSpellOn] = useState(true);
  const spellCache = useRef(new Map<string, boolean>());
  const [spellTick, setSpellTick] = useState(0);
  const [myWords, setMyWords] = useState<Set<string>>(new Set());
  const [ignored, setIgnored] = useState<Set<string>>(new Set());
  const [sp, setSp] = useState<{items: Array<{para: number; start: number; end: number; word: string}>; at: number; suggestions: string[]; replace: string; message: string}>({
    items: [],
    at: 0,
    suggestions: [],
    replace: '',
    message: '',
  });
  /** Insert quote…: the saved quotes, where each source file's details come from, and the flow's state. */
  const [qp, setQp] = useState<{
    mode: 'list' | 'source' | 'zotero' | 'form' | 'preview';
    quotes: SavedQuote[];
    sources: Record<string, SourceRef>;
    chosen: SavedQuote | null;
    source: Source | null;
    page: string;
    doi: string;
    query: string;
    results: Source[];
    form: {type: Details['type']; authors: string; year: string; title: string; container: string; volume: string; issue: string; pages: string; publisher: string; link: string};
    message: string;
    busy: boolean;
  }>({
    mode: 'list',
    quotes: [],
    sources: {},
    chosen: null,
    source: null,
    page: '',
    doi: '',
    query: '',
    results: [],
    form: {type: 'article', authors: '', year: '', title: '', container: '', volume: '', issue: '', pages: '', publisher: '', link: ''},
    message: '',
    busy: false,
  });
  /** Where New puts the document, and the folder being browsed to choose it. */
  const [newFolder, setNewFolder] = useState(DOCUMENTS);
  const [browse, setBrowse] = useState<{path: string; folders?: string[]; error?: string} | null>(null);
  /** Reader text size, as an index into SCALES. */
  const [scaleAt, setScaleAt] = useState(SCALES.indexOf(1));
  const textScale = SCALES[scaleAt];
  const [pageW, setPageW] = useState(0);
  /** The whole document's page starts, counted in the background for one layout (mapKey). */
  /** The page map: counted for one layout (width, page height, text size, fonts) from one version of the text. */
  const [pageMap, setPageMap] = useState<{key: string; layout: string; blocks: Block[]; pages: PageStart[]; done: boolean} | null>(null);
  const [showPages, setShowPages] = useState(false);
  /** The margin column (View ▾); it appears only when the document has changes or comments. */
  const [showMargin, setShowMargin] = useState(true);
  /** Find & replace: what to find, what to put instead, and the match shown last. */
  const [find, setFind] = useState<{query: string; replace: string; matchCase: boolean; at: number}>({query: '', replace: '', matchCase: false, at: -1});
  /** The colour handwritten notes are saved in (remembered). */
  const [inkColor, setInkColor] = useState('#000000');
  /** Fingerprint of the document file as DOCX last read or wrote it; a different one means changed elsewhere. */
  const [diskStamp, setDiskStamp] = useState<string | null>(null);
  /** Unsaved edits found for the document just opened, offered for restoring. */
  const [recovery, setRecovery] = useState<RecoveryRecord | null>(null);
  const [recent, setRecent] = useState<RecentDoc[]>([]);
  const [versions, setVersions] = useState<Array<{path: string; time: number; bytes: number}>>([]);
  /** Documents made with New: saves need no backup or change check on their first save. */
  const [isNew, setIsNew] = useState(false);
  const [pageH, setPageH] = useState(0);
  const [measured, setMeasured] = useState<{key: string; brk: Break} | null>(null);
  const [edits, setEdits] = useState<{steps: Op[][]; cursor: number}>({steps: [], cursor: 0});
  const [saved, setSaved] = useState<{cursor: number; dest: string | null}>({cursor: 0, dest: null});
  const [selection, setSelection] = useState<Selection | null>(null);
  const [caret, setCaret] = useState<Pos | null>(null);
  const [caretBox, setCaretBox] = useState<{key: string; x: number; top: number; height: number} | null>(null);
  const [typing, setTyping] = useState<Typing | null>(null);
  /**
   * Superscript / subscript switched on (or 'none': off) for what is typed next, as in Word:
   * type "H", Subscript, "2", Subscript, "O". Null: typing follows the text before it.
   * Cleared when the caret is moved.
   */
  const [script, setScript] = useState<'sup' | 'sub' | 'none' | null>(null);
  const [input, setInput] = useState('');
  const [discardArmed, setDiscardArmed] = useState(false);
  const started = useRef(false);

  // Ask for write permission before the first log line: refused writes are why the probe
  // lost two logs. Saving needs it too.
  useEffect(() => {
    if (started.current) {
      return;
    }
    started.current = true;
    (async () => {
      // Both up front: reading is needed to list folders for New as well as to open files,
      // and writing for the log and for saving (CT: listing failed because read was never asked).
      setInkOk(await isInkAvailable());
      const canRead = await ensureFileReadPermission();
      const canWrite = await ensureFileWritePermission();
      const name = await Docx?.logName().catch(() => '?');
      const refused = await log(`DOCX opened: NATIVE_BUILD=${nativeBuild()} read=${canRead} write=${canWrite} log=${name}`);
      // Remembered between sessions: text size, New's folder, recent documents.
      try {
        const saved = JSON.parse((await Docx?.load('settings')) ?? '{}') as {scaleAt?: number; newFolder?: string; lastName?: string; author?: string; marginOpen?: boolean; inkColor?: string; citeStyle?: string; spell?: boolean};
        if (typeof saved.lastName === 'string') {
          setLastName(saved.lastName);
        }
        try {
          const z = JSON.parse((await Docx?.load('zotero')) ?? 'null') as ZoteroAccount | null;
          if (z?.userId && z.apiKey) {
            setZotero(z);
          }
        } catch {
          // No account yet.
        }
        if (saved.citeStyle === 'apa' || saved.citeStyle === 'mla' || saved.citeStyle === 'chicago') {
          setCiteStyle(saved.citeStyle);
        }
        if (typeof saved.author === 'string') {
          setAuthor(saved.author);
        }
        if (typeof saved.spell === 'boolean') {
          setSpellOn(saved.spell);
        }
        setMyWords(new Set(parseList<string>(await Docx?.load(`spell-words-${SPELL_LANG}`))));
        if (typeof saved.marginOpen === 'boolean') {
          setShowMargin(saved.marginOpen);
        }
        if (typeof saved.inkColor === 'string') {
          setInkColor(saved.inkColor);
        }
        if (typeof saved.scaleAt === 'number') {
          setScaleAt(Math.max(0, Math.min(SCALES.length - 1, saved.scaleAt)));
        }
        if (saved.newFolder) {
          setNewFolder(saved.newFolder);
        }
        setRecent(JSON.parse((await Docx?.load('recent')) ?? '[]') as RecentDoc[]);
      } catch {
        // First run, or unreadable: defaults.
      }
      if (refused) {
        setStatus(`Log not written: ${refused}`);
      }
    })();
  }, []);

  const applied = useMemo(() => edits.steps.slice(0, edits.cursor).flat(), [edits]);
  // The text after each undo step, kept: a new step applies only its own edits (not every
  // edit since opening), and undo/redo step back to a kept version. Versions share the
  // paragraphs an edit did not touch, so keeping them costs little.
  const undoTexts = useRef<{doc: DocxDocument | null; steps: Op[][]; blocks: Block[][]}>({doc: null, steps: [], blocks: []});
  const committed = useMemo(() => {
    if (!doc) {
      return [];
    }
    const v = undoTexts.current;
    if (v.doc !== doc) {
      undoTexts.current = {doc, steps: [], blocks: [doc.blocks]};
    }
    const cur = undoTexts.current;
    let same = 0;
    while (same < cur.steps.length && same < edits.steps.length && cur.steps[same] === edits.steps[same]) {
      same++;
    }
    cur.steps = cur.steps.slice(0, same);
    cur.blocks = cur.blocks.slice(0, same + 1);
    for (let j = same; j < edits.cursor; j++) {
      cur.blocks.push(applyOps(cur.blocks[j], edits.steps[j]));
      cur.steps.push(edits.steps[j]);
    }
    return cur.blocks[edits.cursor];
  }, [doc, edits]);
  // What is being typed: shown in the page as it is typed, committed when typing ends.
  const pending: Op[] = useMemo(() => {
    const text = typing ? clean(input) : '';
    if (!typing || text === '') {
      return [];
    }
    const r = typedRange(typing);
    const ops: Op[] = [{op: 'text', para: r.para, start: r.start, end: r.end, text}];
    if (script) {
      // 'none' writes baseline: turning superscript off turns both off.
      ops.push({op: 'format', para: r.para, start: r.start, end: r.start + text.length, prop: script === 'none' ? 'sup' : script, on: script !== 'none'});
    }
    return ops;
  }, [typing, input, script]);
  // List numbers are recounted after every edit, as Word does.
  const blocks = useMemo(() => recount(applyOps(committed, pending), doc?.lists ?? {}), [committed, pending, doc]);
  /** Where the caret is drawn: after the typed text while typing. */
  const caretAt: Pos | null = typing
    ? {para: typedRange(typing).para, offset: typedRange(typing).start + clean(input).length}
    : caret;
  const dirty = edits.cursor !== saved.cursor;
  // Documents with tracked changes used to open read-only; they are shown and reviewable now,
  // and every save is verified, so nothing is read-only. Kept as a switch for future cases.
  const readOnly = false;

  // From the committed text: typing never adds blocks, and the window must not change size
  // (redrawing the whole page) as words are typed.
  const counts = useMemo(() => committed.map(wordCount), [committed]);
  const end = useMemo(() => (doc ? windowEnd(counts, anchor.block) : 0), [doc, counts, anchor.block]);
  // While typing, the caret is drawn in the text itself (BlockView), not placed afterwards.
  const window = useMemo(() => {
    const w = blocks.slice(anchor.block, end);
    if (!typing || !caretAt) {
      return w;
    }
    return w.map(b => (b.type === 'p' && b.index === caretAt.para ? {...b, caret: caretAt.offset} : b));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [blocks, anchor.block, end, typing, caretAt?.para, caretAt?.offset]);

  // Measurements belong to one page (window + offset + page height + text size). A new
  // page starts them afresh (the column remounts), and a break computed for another page
  // is never used. Edits keep them: a paragraph an edit changes reports its new layout,
  // and one it doesn't change keeps the layout it reported.
  const pageKey = `${doc?.path}:${anchor.block}:${anchor.offset}:${end}:${pageH}:${textScale}`;
  const measuredFor = useRef('');
  const frames = useRef<Array<Frame | undefined>>([]);
  const lines = useRef<Array<LineBox[] | undefined>>([]);
  const textFrames = useRef<Array<{x: number; y: number} | undefined>>([]);
  const textRefs = useRef<Array<Text | null>>([]);
  if (measuredFor.current !== pageKey) {
    measuredFor.current = pageKey;
    frames.current = [];
    lines.current = [];
    textFrames.current = [];
  }
  const brk: Break = measured?.key === pageKey ? measured.brk : {kind: 'pending'};
  const atEnd = !doc || (brk.kind === 'after' && end >= blocks.length);

  /** A block is measured once it has a frame and, if it is a paragraph with text, its lines. */
  const boxesNow = useCallback(
    (): Array<BlockBox | undefined> =>
      window.map((b, i) => {
        const f = frames.current[i];
        const l = lines.current[i];
        if (!f || (b.type === 'p' && b.runs.length > 0 && !l)) {
          return undefined;
        }
        return {top: f.top, height: f.height, lines: l, forced: b.type === 'p' && !!b.pb};
      }),
    [window],
  );

  const settle = useRef<ReturnType<typeof setTimeout> | null>(null);
  const remeasure = () => {
    if (settle.current) {
      clearTimeout(settle.current);
    }
    const key = pageKey;
    settle.current = setTimeout(() => {
      if (pageH > 0 && measuredFor.current === key) {
        const found = findBreak(boxesNow(), anchor.offset, pageH);
        // Unchanged (most keystrokes): no new state, so no extra redraw.
        setMeasured(m => (m && m.key === key && JSON.stringify(m.brk) === JSON.stringify(found) ? m : {key, brk: found}));
      }
    }, 60);
  };

  // Recomputed when the window changes (an edit): the break is found again from what is measured.
  const remeasureNow = useRef(remeasure);
  remeasureNow.current = remeasure;
  useEffect(() => {
    remeasureNow.current();
  }, [window]);

  // One callback per window slot, kept for the life of the page: paragraphs are drawn again
  // only when they change (BlockView is memoised), so a callback must never go stale.
  const slotCallbacks = useRef<{
    key: string;
    frame: Array<(e: LayoutChangeEvent) => void>;
    lines: Array<(l: LineBox[]) => void>;
    text: Array<(e: LayoutChangeEvent) => void>;
    ref: Array<(t: Text | null) => void>;
  }>({key: '', frame: [], lines: [], text: [], ref: []});
  if (slotCallbacks.current.key !== pageKey) {
    slotCallbacks.current = {key: pageKey, frame: [], lines: [], text: [], ref: []};
  }
  /** The slot's callback of one kind, made once per page. */
  const slot = <T,>(list: T[], i: number, make: () => T): T => {
    if (list[i] === undefined) {
      list[i] = make();
    }
    return list[i];
  };

  const onFrame = (i: number) =>
    slot(slotCallbacks.current.frame, i, () => (e: LayoutChangeEvent) => {
      const {x, y, height} = e.nativeEvent.layout;
      frames.current[i] = {x, top: y, height};
      remeasureNow.current();
    });

  const onLines = (i: number) =>
    slot(slotCallbacks.current.lines, i, () => (l: LineBox[]) => {
      lines.current[i] = l;
      remeasureNow.current();
    });

  const onTextFrame = (i: number) =>
    slot(slotCallbacks.current.text, i, () => (e: LayoutChangeEvent) => {
      const {x, y} = e.nativeEvent.layout;
      textFrames.current[i] = {x, y};
    });

  const textRefFor = (i: number) =>
    slot(slotCallbacks.current.ref, i, () => (t: Text | null) => {
      textRefs.current[i] = t;
    });

  // ---------------------------------------------------------------- open / close

  /** Unsaved changes need a second tap to be discarded (Open and New). */
  const mayDiscard = (action: string): boolean => {
    flushTyping();
    if ((dirty || pending.length > 0) && !discardArmed) {
      setDiscardArmed(true);
      setStatus(`You have unsaved changes. Tap ${action} again to discard them, or Save first.`);
      return false;
    }
    setDiscardArmed(false);
    return true;
  };

  /** Opens a document and makes its fonts available. `extra` marks a document made with New. */
  /**
   * Opens a document. Edits always apply to a private snapshot of it as opened (`source`)
   * and are saved over the document itself (`saveTo`), so saving again and again stays right.
   * `created` marks a document just made with New (its blank is already the snapshot).
   */
  const openPath = async (path: string, created?: {source: string}) => {
    setStatus('Opening…');
    const loaded = await Docx!.open(path);
    const key = docKey(path);
    const source = created?.source ?? (await Docx!.snapshot(path, key));
    const stamp = await Docx!.fileStamp(path);
    const opened: DocxDocument = {...loaded, source, saveTo: path};
    setDoc(opened);
    setIsNew(!!created);
    setDiskStamp(stamp);
    setPad(null);
    setIgnored(new Set());
    const found = parseRecovery(await Docx!.load(`recovery-${key}`), path, stamp);
    setRecovery(found);
    setRecent(list => {
      const next = touchRecent(list, path, opened.name);
      Docx?.store('recent', JSON.stringify(next));
      return next;
    });
    setAnchor(START);
    setHistory([]);
    setContents(false);
    setShowPages(false);
    setPageMap(null);
    setMenu(null);
    setEdits({steps: [], cursor: 0});
    setSaved({cursor: 0, dest: null});
    setSelection(null);
    setCaret(null);
    setTyping(null);
    const r = opened.report;
    setStatus(
      r.trackedChanges > 0
        ? 'This document has tracked changes: they are marked in the text and listed in the margin.'
        : created
        ? 'New document. Tap the page to start typing.'
        : '',
    );
    if (found) {
      setMenu('recover');
    }
    log(`opened ${opened.name}: ${opened.blocks.length} blocks in ${opened.ms} ms; ${JSON.stringify(r)}`);
    try {
      setFonts(new Set(await Docx!.fonts(withStandIns(fontsUsed(opened.blocks)))));
    } catch (error) {
      log(`fonts failed: ${errorText(error)}`);
    }
  };

  const open = async () => {
    if (!mayDiscard('Open')) {
      return;
    }
    setBusy(true);
    try {
      if (!(await ensureFileReadPermission())) {
        setStatus('File access was not allowed.');
        return;
      }
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0, // 1 opens the file and never returns (reference_sdk_calls_that_navigate)
        maxNum: 1,
        title: 'Open a Word document',
        rightButtonText: 'Open',
        suffixList: ['docx'],
      })) as string[] | null | undefined;
      const path = picked?.find(p => typeof p === 'string' && p.length > 0);
      if (!path) {
        return;
      }
      if (!/\.docx$/i.test(path)) {
        setStatus('Only .docx files can be opened (not .doc).');
        return;
      }
      await openPath(path);
    } catch (error) {
      setStatus(`Could not open: ${errorText(error)}`);
      log(`open failed: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  const startNew = () => {
    if (!mayDiscard('New')) {
      return;
    }
    setTemplate(null);
    setNaming('Untitled');
    setMenu('name');
  };

  /** New from template: pick a .docx or .dotx; the new document starts as a copy of it. */
  const startFromTemplate = async () => {
    if (!mayDiscard('New')) {
      return;
    }
    setMenu(null);
    try {
      if (!(await ensureFileReadPermission())) {
        setStatus('File access was not allowed.');
        return;
      }
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0,
        maxNum: 1,
        title: 'Choose a template or document to start from',
        rightButtonText: 'Choose',
        suffixList: ['docx', 'dotx'],
      })) as string[] | null | undefined;
      const path = picked?.find(p => typeof p === 'string' && p.length > 0);
      if (!path) {
        return;
      }
      if (!/\.(docx|dotx)$/i.test(path)) {
        setStatus('Choose a .docx or .dotx file.');
        return;
      }
      setTemplate(path);
      setNaming(path.slice(path.lastIndexOf('/') + 1).replace(/\.(docx|dotx)$/i, ''));
      setMenu('name');
    } catch (error) {
      setStatus(`No template chosen: ${errorText(error)}`);
    }
  };

  /** A blank document in the Document folder, opened for editing; it saves over itself. */
  const createNew = async () => {
    const name = (naming ?? '').trim() || 'Untitled';
    const from = template;
    setNaming(null);
    setTemplate(null);
    setMenu(null);
    setBusy(true);
    try {
      if (!(await ensureFileWritePermission())) {
        setStatus('Making a document needs file write permission.');
        return;
      }
      const made = from ? await Docx!.createFrom(from, name, newFolder) : await Docx!.create(name, newFolder);
      await openPath(made.path, {source: made.source});
    } catch (error) {
      setStatus(`Could not make the document: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  // ---------------------------------------------------------------- fonts and sizes

  const applyRunStyle = (style: {font?: string; size?: number}, label: string) => {
    const ranges = targets();
    setMenu(null);
    if (ranges.length === 0) {
      nothingSelected();
      return;
    }
    commit(ranges.map(r => ({op: 'runStyle', para: r.para, start: r.start, end: r.end, ...style})), label);
  };

  /** A font file picked by hand: named from inside the file, loaded, and remembered. */
  const addFont = async () => {
    try {
      if (!(await ensureFileReadPermission())) {
        setStatus('File access was not allowed.');
        return;
      }
      // Several files at once: a family's regular, bold, italic … in one go.
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0,
        maxNum: 24,
        title: 'Choose font files (.ttf or .otf)',
        rightButtonText: 'Add',
        suffixList: ['ttf', 'otf'],
      })) as string[] | null | undefined;
      const paths = (picked ?? []).filter(p => typeof p === 'string' && /\.(ttf|otf)$/i.test(p));
      const families = new Set<string>();
      const failed: string[] = [];
      for (const path of paths) {
        try {
          families.add((await Docx!.addFont(path)).family);
        } catch {
          failed.push(path.split('/').pop() ?? path);
        }
      }
      if (families.size > 0) {
        setFonts(f => new Set([...f, ...families]));
      }
      setStatus(
        [families.size ? `Added ${[...families].join(', ')}.` : '', failed.length ? `Couldn't read ${failed.join(', ')}.` : '']
          .filter(Boolean)
          .join(' ') || 'No font files chosen.',
      );
    } catch (error) {
      setStatus(`Fonts not added: ${errorText(error)}`);
    }
  };

  /**
   * The font and size the Font and Size menus tick: of the selection's first character, or
   * the character before the caret (what typed text takes); text without its own takes the
   * paragraph's base (its style and the document's defaults).
   */
  const current = (): {font?: string; size?: number; hc?: string; s?: boolean; sup?: boolean; sub?: boolean} => {
    const sel = !typing && selection ? selection : null;
    const at = sel ? (comparePos(sel.from, sel.to) <= 0 ? sel.from : sel.to) : caretAt;
    const p = at ? paragraph(at.para) : undefined;
    if (!p || !at) {
      return {};
    }
    let offset = 0;
    let run: Run | undefined = p.runs[0];
    for (const r of p.runs) {
      const end = offset + r.t.length;
      if (sel ? at.offset >= offset && at.offset < end : at.offset > offset && at.offset <= end) {
        run = r;
        break;
      }
      offset = end;
    }
    return {font: run?.f ?? p.bf, size: run?.sz ?? p.bs, hc: run?.h ? run.hc ?? 'yellow' : undefined, s: run?.s, sup: run?.sup, sub: run?.sub};
  };

  const fontChoices = useMemo(() => {
    const bases = blocks.flatMap(b => (b.type === 'p' && b.bf ? [b.bf] : []));
    return [...new Set([...fonts, ...fontsUsed(blocks), ...bases])].sort((a, b) => a.localeCompare(b));
  }, [fonts, blocks]);
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

  const comments = useMemo(() => commentsAfter(doc?.comments, applied), [doc, applied]);
  // From the committed text: typing never adds changes or notes, and this runs on every key.
  const hasChanges = useMemo(() => committed.some(b => b.type === 'p' && !!b.revs?.length), [committed]);
  const inks = useMemo(() => inksAfter(doc?.inks, applied), [doc, applied]);
  const hasNotes = useMemo(
    () => hasChanges || committed.some(b => b.type === 'p' && (!!b.marks?.length || b.runs.some(r => r.obj === 'ink'))),
    [hasChanges, committed],
  );
  // The margin column is there whenever the document has notes, comments or changes: open,
  // or folded into a narrow strip.
  const marginOn = hasNotes && !showPages && !contents;
  // The reading page's text column: the page less the margin column. Pages are counted at
  // this width even while Pages or Contents shows (they hide the margin column), so a page
  // number means the same page everywhere.
  const readW = hasNotes ? Math.max(200, pageW - (showMargin ? MARGIN_W : MARGIN_STRIP) - MARGIN_GAP) : pageW;
  const textW = marginOn ? readW : pageW;

  // The page map belongs to one layout (width, page height, text size, fonts) and one
  // version of the committed text. It is recounted a moment after an edit settles — from
  // the page the edit is on — and not while typing: typing is committed first.
  const layoutKey = `${doc?.path}:${textScale}:${readW}x${pageH}:${fonts.size}`;
  const versionNo = useRef(0);
  // A new number for each version of the committed text.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const textVersion = useMemo(() => ++versionNo.current, [committed]);
  const mapKey = `${layoutKey}:${textVersion}`;
  const [counting, setCounting] = useState<{key: string; layout: string; blocks: Block[]; seed?: PageStart[]} | null>(null);
  const pageMapNow = useRef(pageMap);
  pageMapNow.current = pageMap;
  useEffect(() => {
    if (!doc || pageH <= 0 || pageW <= 0) {
      return;
    }
    const t = setTimeout(() => {
      const prev = pageMapNow.current;
      let seed: PageStart[] | undefined;
      if (prev && prev.done && prev.layout === layoutKey) {
        const d = firstDifference(prev.blocks, committed);
        if (d > 0) {
          seed = prev.pages.slice(0, pageIndexOf(prev.pages, {block: d, offset: 0}) + 1);
        }
      }
      setCounting({key: mapKey, layout: layoutKey, blocks: committed, seed});
    }, 1500);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mapKey, doc, pageH, pageW]);
  // A recount keeps showing the last full count until it is done (no page number jumping on e-ink).
  const map = pageMap && pageMap.layout === layoutKey ? pageMap : null;
  const pageNumber = map ? pageIndexOf(map.pages, anchor) + 1 : null;
  // From the committed text (counts are remembered per paragraph): for the target on the bottom line.
  const words = useMemo(() => wordCounts(committed), [committed]);

  // The target belongs to the document.
  useEffect(() => {
    setTarget(null);
    setTargetText('');
    if (!doc) {
      return;
    }
    (async () => {
      try {
        const n = Number(JSON.parse((await Docx?.load(`target-${docKey(doc.path)}`)) ?? 'null'));
        if (Number.isInteger(n) && n > 0) {
          setTarget(n);
          setTargetText(String(n));
        }
      } catch {
        // None set.
      }
    })();
  }, [doc?.path]);

  const saveTarget = (text: string) => {
    if (!doc) {
      return;
    }
    const n = Number(text);
    const value = Number.isInteger(n) && n > 0 ? n : null;
    setTarget(value);
    setTargetText(value ? String(value) : '');
    Docx?.store(`target-${docKey(doc.path)}`, JSON.stringify(value));
    setStatus(value ? `Target set: ${value.toLocaleString()} words of body text.` : 'Target removed.');
  };

  // Only while the Pages view shows: what each page holds, for its card.
  const pageInfos = useMemo(() => {
    if (!showPages || !map) {
      return [];
    }
    const replies = new Set(comments.filter(c => c.parent !== undefined).map(c => c.id));
    return map.pages.map((_, i) => pageInfo(map.blocks, map.pages, i, id => replies.has(id)));
  }, [showPages, map, comments]);

  /** For checking page numbers against the screen: the count, and the page each heading starts on. */
  const logPages = (bs: Block[], pages: PageStart[]) => {
    const heads = outline(bs)
      .slice(0, 40)
      .map(h => `"${h.text.slice(0, 30)}" p${pageIndexOf(pages, {block: h.block, offset: 0}) + 1}`);
    log(`pages: ${pages.length} at ${readW}x${pageH} scale ${textScale}; headings ${heads.join(', ') || 'none'}`);
  };

  /** A page start from the page map: the page `at` is on. Without a count reaching it, `at` itself. */
  const pageStartOf = (at: Anchor): Anchor => {
    if (!map || map.pages.length === 0) {
      return at;
    }
    const i = pageIndexOf(map.pages, at);
    if (!map.done && i === map.pages.length - 1) {
      return at; // the count has not got past it yet
    }
    return map.pages[i].anchor;
  };

  // Which page shows, in the log: to compare with the Pages view when they disagree.
  useEffect(() => {
    if (doc && pageNumber !== null) {
      const b = blocks[anchor.block];
      const text = b?.type === 'p' ? paragraphText(b).slice(0, 30) : b?.type ?? '';
      log(`showing page ${pageNumber}${map?.done ? '' : '?'} (block ${anchor.block}+${anchor.offset}: "${text}")`);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [anchor.block, anchor.offset]);

  /** Turns to the page where character `offset` of paragraph `para` is (a find match, a misspelling). */
  const showPlace = (para: number, offset: number) => {
    const block = blocks.findIndex(b => b.type === 'p' && b.index === para);
    const b = blocks[block];
    // A long paragraph can run across pages: the page is found by the character.
    const page = map && b?.type === 'p' ? pageOfChar(map.pages, block, toShown(b, offset), map.done) : -1;
    if (page >= 0) {
      const at = map!.pages[page].anchor;
      if (at.block !== anchor.block || at.offset !== anchor.offset) {
        goTo(at);
      }
    } else if (block >= 0 && (block < anchor.block || block >= end)) {
      goTo({block, offset: 0});
    }
  };

  /** Goes to the page `at` is on (◀ and ▶ then step page by page from there). */
  const goTo = (at: Anchor) => {
    flushTyping();
    setHistory([]);
    setAnchor(pageStartOf(at));
    setShowPages(false);
    setContents(false);
  };

  const goToEnd = () => {
    if (map?.done) {
      goTo(map.pages[map.pages.length - 1].anchor);
    } else if (blocks.length > 0) {
      goTo({block: blocks.length - 1, offset: 0});
    }
  };

  /** A new text size re-flows the pages; the page starts again at the top of its first block. */
  const setScale = (i: number) => {
    const next = Math.max(0, Math.min(SCALES.length - 1, i));
    if (next !== scaleAt) {
      setScaleAt(next);
      setAnchor(a => ({block: a.block, offset: 0}));
      setHistory([]);
    }
  };

  // After the page size, text size or margin column changes, the page shown is re-aligned
  // to a page start once the new count reaches it, so ◀ ▶ and the page number agree.
  const alignedFor = useRef('');
  useEffect(() => {
    if (!map || alignedFor.current === map.layout || !doc) {
      return;
    }
    const at = pageStartOf(anchor);
    if (at === anchor && !map.done) {
      return; // not counted this far yet
    }
    alignedFor.current = map.layout;
    if (at.block !== anchor.block || at.offset !== anchor.offset) {
      setAnchor(at);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [map]);

  /** Where a page begins, as a paragraph position (a table's page begins at the next paragraph). */
  const pageStartPos = (pg: PageStart): Pos | null => {
    for (let i = pg.anchor.block; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'p') {
        // The page map counts what the screen shows (indent spacer, deleted text) as characters.
        const offset = i === pg.anchor.block && pg.char > 0 ? fromShown(b, pg.char) : 0;
        return {para: b.index, offset: Math.min(offset, paragraphText(b).length)};
      }
    }
    return null;
  };

  const lastParagraph = (): ParagraphBlock | undefined =>
    [...blocks].reverse().find((b): b is ParagraphBlock => b.type === 'p');

  /** A blank page after page i: an empty paragraph that starts a new page, and what followed starts another. */
  const newPageAfter = (i: number) => {
    if (!map) {
      return;
    }
    flushTyping();
    const next = map.pages[i + 1];
    const ops: Op[] = [];
    let blank: Pos;
    if (next) {
      const at = pageStartPos(next);
      if (!at) {
        return;
      }
      const p = paragraph(at.para);
      const problem = p ? splitProblem(p, at.offset) : 'Paragraph not found.';
      if (problem) {
        setStatus(problem);
        return;
      }
      if (at.offset > 0) {
        ops.push({op: 'split', para: at.para, offset: at.offset});
        at.para += 1;
      }
      // An empty paragraph in front of the next page's text, both starting new pages.
      ops.push({op: 'split', para: at.para, offset: 0});
      ops.push({op: 'para', para: at.para, pb: true});
      ops.push({op: 'para', para: at.para + 1, pb: true});
      blank = {para: at.para, offset: 0};
    } else {
      const last = lastParagraph();
      if (!last) {
        return;
      }
      ops.push({op: 'split', para: last.index, offset: paragraphText(last).length});
      ops.push({op: 'para', para: last.index + 1, pb: true});
      blank = {para: last.index + 1, offset: 0};
    }
    commit(ops, 'new page');
    setShowPages(false);
    setSelection(null);
    setCaret(blank);
    keys.current?.focus();
    setStatus('New page. Type to fill it.');
  };

  /**
   * Page card buttons take a second tap: a finger scrolling the Pages view can end on one
   * (CT got a blank page added in the middle of a paper that way).
   */
  const [armed, setArmed] = useState<string | null>(null);
  useEffect(() => {
    if (!armed) {
      return;
    }
    const t = setTimeout(() => setArmed(null), 6000);
    return () => clearTimeout(t);
  }, [armed]);

  const cardButton = (key: string, label: string, question: string, action: () => void) => (
    <Pressable
      key={key}
      onPress={once(key, () => {
        if (armed === key) {
          setArmed(null);
          action();
        } else {
          setArmed(key);
          setStatus(question);
        }
      })}
      style={[styles.pageCardButton, armed === key ? styles.pageCardButtonArmed : null]}>
      <Text allowFontScaling={false} style={[styles.pageCardButtonText, armed === key ? styles.pageCardButtonTextArmed : null]}>
        {armed === key ? 'Tap again' : label}
      </Text>
    </Pressable>
  );

  /** The edit that removes the page break page i starts with, or null when it starts with none. */
  const pageBreakAt = (i: number): Op[] | null => {
    const at = map ? pageStartPos(map.pages[i]) : null;
    return at ? pageBreakRemoval(at) : null;
  };

  const removePageBreak = (i: number) => {
    const ops = pageBreakAt(i);
    if (!ops) {
      return;
    }
    flushTyping();
    commit(ops, 'remove page break');
    setStatus(`Removed the page break before page ${i + 1}. Undo puts it back.`);
  };

  /** Selects exactly the text on page i, to see it before deleting it. */
  const selectPage = (i: number) => {
    if (!map) {
      return;
    }
    flushTyping();
    const from = pageStartPos(map.pages[i]);
    const next = map.pages[i + 1];
    const last = lastParagraph();
    const to = next ? pageStartPos(next) : last ? {para: last.index, offset: paragraphText(last).length} : null;
    if (!from || !to || comparePos(from, to) >= 0) {
      setStatus('That page has no text to select.');
      return;
    }
    goTo(map.pages[i].anchor);
    setCaret(null);
    setSelection({from, to});
    keys.current?.focus();
    setStatus('Page selected. Delete (or Backspace) removes it; tap the page to cancel.');
  };

  // ---------------------------------------------------------------- folder for New

  const shortPath = (path: string) => path.replace(`${STORAGE}/`, '').replace(STORAGE, 'Internal storage');

  const browseTo = async (path: string) => {
    setBrowse({path});
    setMenu('folder');
    if (!(await ensureFileReadPermission())) {
      setBrowse({path, error: 'File access was not allowed.'});
      return;
    }
    const r = await Docx!.listFolders(path);
    setBrowse({path, folders: r.folders, error: r.error});
  };

  /** Fallback when a folder can't be listed: the folder of any file picked in it. */
  const folderFromFile = async () => {
    try {
      // The same call Open makes. No starting folder: needSelectFolder opens a different,
      // non-Ratta view that refuses paths outside its whitelist (CT; SNFolio saw the same).
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0,
        maxNum: 1,
        title: 'Choose any file in the folder you want',
        rightButtonText: 'Choose',
      })) as string[] | null | undefined;
      const path = picked?.find(p => typeof p === 'string' && p.includes('/'));
      if (path) {
        setNewFolder(path.slice(0, path.lastIndexOf('/')));
      }
    } catch (error) {
      setStatus(`No folder chosen: ${errorText(error)}`);
    }
    setBrowse(null);
    setMenu('name');
  };

  const close = () => {
    if (pad) {
      setPad(null);
    }
    // The pen engine must never outlive its pad (PluginHost wedges).
    deactivateInk().finally(() => PluginManager.closePluginView());
  };

  // ---------------------------------------------------------------- paging

  const next = () => {
    if (!doc || atEnd) {
      return;
    }
    flushTyping();
    const to = anchorAfter(anchor, boxesNow(), brk);
    if (!to || to.block >= blocks.length) {
      return;
    }
    setHistory(h => [...h, anchor]);
    setAnchor(to);
  };

  const atStart = anchor.block === 0 && anchor.offset === 0;

  // ---------------------------------------------------------------- pictures and figures

  /** Insert picture…: after the paragraph with the caret (or the end of the selection). */
  const startPicture = () => {
    const at = !typing && selection ? (comparePos(selection.from, selection.to) <= 0 ? selection.to : selection.from) : caretAt;
    flushTyping();
    if (!at) {
      setStatus('Tap the paragraph the picture goes after.');
      return;
    }
    setPic(p => ({...p, para: at.para, path: undefined, w: undefined, h: undefined, title: '', note: ''}));
    setMenu('picture');
  };

  const choosePicture = async () => {
    try {
      if (!(await ensureFileReadPermission())) {
        setPic(p => ({...p, note: 'File access was not allowed.'}));
        return;
      }
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0,
        maxNum: 1,
        title: 'Choose a picture (PNG or JPEG)',
        rightButtonText: 'Choose',
        suffixList: ['png', 'jpg', 'jpeg', 'gif', 'bmp'],
      })) as string[] | null | undefined;
      const path = picked?.find(x => typeof x === 'string' && /\.(png|jpe?g|gif|bmp)$/i.test(x));
      if (!path) {
        return;
      }
      Image.getSize(
        `file://${path}`,
        (w, h) => setPic(p => ({...p, path, w, h, note: `${path.slice(path.lastIndexOf('/') + 1)} · ${w} × ${h} pixels`})),
        () => setPic(p => ({...p, path: undefined, note: 'That picture could not be read.'})),
      );
    } catch (error) {
      setPic(p => ({...p, note: `No picture chosen: ${errorText(error)}`}));
    }
  };

  const insertPicture = () => {
    const para = paragraph(pic.para);
    if (!doc || !para || !pic.path || !pic.w || !pic.h) {
      return;
    }
    const problem = splitProblem(para, paragraphText(para).length);
    if (problem) {
      setPic(p => ({...p, note: problem}));
      return;
    }
    // As big as it is (at 96 dpi), no wider than the text and no taller than two-thirds of the page.
    const page = pageAfter(doc.page, applied);
    const twip = 635;
    const maxW = Math.max(1440, page.width - page.left - page.right) * twip;
    const maxH = Math.max(1440, Math.round(((page.height - page.top - page.bottom) * 2) / 3)) * twip;
    const {cx, cy} = pictureEmu(pic.w, pic.h, maxW, maxH);
    const alt = pic.label && pic.title.trim() ? pic.title.trim() : pic.path.slice(pic.path.lastIndexOf('/') + 1);
    const {ops, n} = figureOps(blocks, para, {path: pic.path, cx, cy, alt}, pic.label ? pic.title : null, citeStyle);
    setMenu(null);
    commit(ops, 'insert picture');
    setSelection(null);
    setCaret(null);
    const name = CITE_STYLES.find(x => x.id === citeStyle)?.name ?? '';
    setStatus(n ? `Figure ${n} inserted (${name} style). Later figures were renumbered; check mentions like “see Figure ${n}” in the text.` : 'Picture inserted.');
  };

  /** Go to page…: from View, or by tapping the page number on the bottom line. */
  const openGoTo = () => {
    flushTyping();
    setGoToText('');
    setGoToNote('');
    setMenuX(Math.max(0, pageW - MENU_W));
    setMenu('goto');
  };

  const goToPage = () => {
    const n = Number(goToText);
    if (!map) {
      setGoToNote('Still counting pages. Try again in a moment.');
      return;
    }
    if (!Number.isInteger(n) || n < 1 || n > map.pages.length) {
      setGoToNote(map.done ? `There are ${map.pages.length} pages.` : `Counted ${map.pages.length} pages so far.`);
      return;
    }
    setMenu(null);
    goTo(map.pages[n - 1].anchor);
  };

  /**
   * ◀: the page before this one, as in any reader. From the page map; a page shown from
   * part-way down (before the map counted it) goes to the start of its own page first.
   * Before the count reaches here, back through the pages turned with ▶.
   */
  const previous = () => {
    if (!doc || atStart) {
      return;
    }
    flushTyping();
    if (map && map.pages.length > 0) {
      const i = pageIndexOf(map.pages, anchor);
      const counted = map.done || i < map.pages.length - 1;
      if (counted) {
        const start = map.pages[i].anchor;
        const onStart = start.block === anchor.block && Math.abs(start.offset - anchor.offset) <= 1;
        const to = onStart ? map.pages[i - 1]?.anchor : start;
        if (to) {
          setHistory([]);
          setAnchor(to);
        }
        return;
      }
    }
    if (history.length > 0) {
      setAnchor(history[history.length - 1]);
      setHistory(history.slice(0, -1));
    } else {
      setStatus('Still counting pages. ◀ works in a moment.');
    }
  };

  /** From Contents: the page the heading is on. */
  const jump = (block: number) => {
    flushTyping();
    setHistory([]);
    setAnchor(pageStartOf({block, offset: 0}));
    setContents(false);
  };

  // ---------------------------------------------------------------- pen selection

  type Hit = {para: number; offset: number; char: number};

  /** The paragraph under a viewport point (by frame), then the character (natively). */
  const hit = async (x: number, y: number): Promise<Hit | string> => {
    const colY = y + anchor.offset;
    let best: {i: number; d: number} | null = null;
    window.forEach((b, i) => {
      const f = frames.current[i];
      if (b.type !== 'p' || !f) {
        return;
      }
      const d = colY < f.top ? f.top - colY : colY > f.top + f.height ? colY - (f.top + f.height) : 0;
      if (!best || d < best.d) {
        best = {i, d};
      }
    });
    if (!best) {
      return 'no paragraph on this page';
    }
    const i = (best as {i: number}).i;
    const p = window[i] as ParagraphBlock;
    const f = frames.current[i]!;
    const tf = textFrames.current[i] ?? {x: 0, y: 0};
    const tag = findNodeHandle(textRefs.current[i] ?? null);
    if (!DocxText || tag === null) {
      return 'text lookup unavailable';
    }
    if (p.runs.length === 0) {
      return {para: p.index, offset: 0, char: 0};
    }
    const scale = PixelRatio.get();
    const lx = Math.max(0, x - f.x - tf.x) * scale;
    const ly = Math.max(0, Math.min(f.height - 1, colY - f.top - tf.y)) * scale;
    const res = await DocxText.offsetAt(tag, lx, ly);
    if (res.error !== undefined) {
      return `${res.error} (${res.viewClass})`;
    }
    // The TextView also counts the indent spacer and shown deleted text; the model does not.
    return {para: p.index, offset: fromShown(p, res.offset!), char: fromShown(p, res.char!)};
  };

  const textOf = (para: number): string => {
    const b = blocks.find(x => x.type === 'p' && x.index === para) as ParagraphBlock | undefined;
    return b ? paragraphText(b) : '';
  };

  const penFrom = useRef<{x: number; y: number} | null>(null);
  const penTo = useRef<{x: number; y: number} | null>(null);
  const selecting = useRef(false);
  const lastTap = useRef<{x: number; y: number; at: number} | null>(null);
  /** Pen and finger, each logged the first time it touches the page. */
  const toolsSeen = useRef(new Set<string>());
  /** The invisible field that receives the keyboard. Focused on the first tap, kept until Hide keyboard. */
  const keys = useRef<TextInput>(null);

  const release = async () => {
    const a = penFrom.current;
    const b = penTo.current;
    penFrom.current = null;
    penTo.current = null;
    if (!a || !b || selecting.current || readOnly) {
      return;
    }
    selecting.current = true;
    try {
      // A finger only turns pages: swipe left for the next page, right for the one before.
      // The pen does everything else. (Where the native layer is missing, all is pen.)
      const tool = DocxText?.gestureTool ? await DocxText.gestureTool().catch(() => null) : null;
      if (tool && !toolsSeen.current.has(tool.tool)) {
        toolsSeen.current.add(tool.tool);
        log(`touch: first ${tool.tool} (Android tool type ${tool.toolType})`);
      }
      if (tool?.tool === 'finger') {
        const dx = b.x - a.x;
        const dy = b.y - a.y;
        if (Math.abs(dx) >= SWIPE_MIN && Math.abs(dx) > Math.abs(dy) * 1.5) {
          if (dx < 0) {
            next();
          } else {
            previous();
          }
        }
        return;
      }
      const tap = Math.hypot(b.x - a.x, b.y - a.y) < TAP_SLOP;
      const now = Date.now();
      const prev = lastTap.current;
      const doubleTap = tap && !!prev && now - prev.at < DOUBLE_TAP_MS && Math.hypot(a.x - prev.x, a.y - prev.y) < TAP_SLOP * 2;
      lastTap.current = tap && !doubleTap ? {x: a.x, y: a.y, at: now} : null;
      const ha = await hit(a.x, a.y);
      const hb = tap ? ha : await hit(b.x, b.y);
      if (typeof ha === 'string' || typeof hb === 'string') {
        setStatus(typeof ha === 'string' ? ha : (hb as string));
        return;
      }
      // Anything typed so far is committed first. The page already showed it, so the
      // positions just measured stay right.
      flushTyping();
      keys.current?.focus();
      setScript(null);
      if (tap && !doubleTap) {
        // One tap: a caret in the gap nearest the pen. The next key types there.
        setSelection(null);
        setCaret({para: ha.para, offset: ha.offset});
        setStatus('');
        return;
      }
      // Order by the character under each end. A double tap takes the word; a drag takes
      // whole words, or exact characters when it stays inside one word.
      const [s, e] = comparePos({para: ha.para, offset: ha.char}, {para: hb.para, offset: hb.char}) <= 0 ? [ha, hb] : [hb, ha];
      const ws = doubleTap ? wordAround(textOf(s.para), s.char, 'right') : null;
      const picked = doubleTap
        ? ws
          ? {from: {para: s.para, offset: ws.start}, to: {para: s.para, offset: ws.end}}
          : null
        : penSelection(textOf(s.para), textOf(e.para), s, e);
      if (!picked) {
        setSelection(null);
        return;
      }
      setSelection(picked);
      setCaret(null);
      setTyping(null);
      setStatus('');
    } catch (error) {
      setStatus(`Selection failed: ${errorText(error)}`);
      log(`selection failed: ${errorText(error)}`);
    } finally {
      selecting.current = false;
    }
  };

  // The touch layer starts PAD left of the page's text column; points are in page coordinates.
  const point = (e: GestureResponderEvent) => ({x: e.nativeEvent.locationX - PAD, y: e.nativeEvent.locationY});

  const selectedIn = (b: (typeof window)[number]): {start: number; end: number} | null => {
    // Once replacement text is typed, the selected text is gone from the page.
    if (!selection || b.type !== 'p' || pending.length > 0) {
      return null;
    }
    const r = rangesBetween([b], selection.from, selection.to)[0];
    return r ? {start: r.start, end: r.end} : null;
  };

  // ---------------------------------------------------------------- edits

  /** One button press = one undo step, however many paragraphs it touches. */
  const commit = (ops: Op[], label: string) => {
    if (ops.length === 0) {
      return;
    }
    setEdits(({steps, cursor}) => ({steps: [...steps.slice(0, cursor), ops], cursor: cursor + 1}));
    // An undone save point can no longer be reached by redo.
    setSaved(sv => (sv.cursor > edits.cursor ? {...sv, cursor: -1} : sv));
    log(`${label}: ${JSON.stringify(ops)}`);
  };

  /**
   * What a tool acts on: the selection, or else the word at the caret (as in Word). Typing
   * in progress is committed first; the page already showed it, so positions stay right.
   */
  const targets = (): Range[] => {
    if (!typing && selection) {
      return rangesBetween(blocks, selection.from, selection.to);
    }
    const at = caretAt;
    flushTyping();
    if (!at) {
      return [];
    }
    const text = textOf(at.para);
    const ch = at.offset < text.length && /\S/.test(text[at.offset]) ? at.offset : at.offset - 1;
    const w = ch >= 0 ? wordAround(text, ch, 'left') : null;
    return w && w.end > at.offset - 1 && w.start <= at.offset ? [{para: at.para, start: w.start, end: w.end}] : [];
  };

  /** Paragraph-level tools act on the paragraphs of the selection, or the caret's paragraph. */
  const targetParagraphs = (): number[] => {
    if (!typing && selection) {
      const [s0, e0] = comparePos(selection.from, selection.to) <= 0 ? [selection.from, selection.to] : [selection.to, selection.from];
      return blocks.filter(b => b.type === 'p' && b.index >= s0.para && b.index <= e0.para).map(b => (b as ParagraphBlock).index);
    }
    const at = caretAt;
    flushTyping();
    return at ? [at.para] : [];
  };

  const nothingSelected = () => setStatus('Select some text first, or tap inside a word.');

  // ---------------------------------------------------------------- handwritten notes

  /**
   * Handwritten note…: a narrow pen pad; the ink goes into the Word file as a picture in the
   * right margin beside the selected words (or the caret), where Word shows and prints it.
   */
  const startNote = () => {
    setMenu(null);
    const ranges = !typing && selection ? rangesBetween(blocks, selection.from, selection.to).filter(r => r.end > r.start) : [];
    let at: {para: number; at: number; quote: string} | null = null;
    if (ranges.length > 0) {
      const r = ranges[0];
      at = {para: r.para, at: r.start, quote: textOf(r.para).slice(r.start, r.end).slice(0, 200)};
    } else if (caretAt) {
      flushTyping();
      const w = targets()[0];
      at = w ? {para: w.para, at: w.start, quote: textOf(w.para).slice(w.start, w.end)} : {para: caretAt.para, at: caretAt.offset, quote: ''};
    }
    if (!at) {
      setStatus('Select the words the note is about, or tap where it goes.');
      return;
    }
    const p = paragraph(at.para);
    const problem = p ? textEditProblem(p, at.at, at.at) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    done();
    // After the keyboard has gone: its leaving resizes the page, and the pen pad must
    // never move or resize once the engine is bound to it.
    const target = at;
    setTimeout(() => setPad(target), 400);
  };

  // The pen engine binds once the pad's surface is on screen, and is released when it goes.
  useEffect(() => {
    if (!pad) {
      return;
    }
    let live = true;
    activateInk().then(r => {
      if (live && r !== 'active') {
        setStatus(`The pen pad could not start (${r}).`);
        log(`ink activate: ${r}`);
      }
    });
    return () => {
      live = false;
      deactivateInk();
    };
    // Only when a pad opens or closes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pad !== null]);

  /** Keep: the ink becomes a picture (private storage until saved) and an edit that places it. */
  const saveNote = async () => {
    if (!pad || !doc || !DocxInk) {
      return;
    }
    try {
      const dir = await DocxInk.notesDir(docKey(doc.saveTo ?? doc.path));
      const id = Date.now().toString(36);
      const res = await DocxInk.save(`${dir}/${id}.png`, inkColor);
      if (res.empty || !res.path) {
        setStatus('Write the note first, or Cancel.');
        return;
      }
      commit([{op: 'ink', para: pad.para, at: pad.at, id, png: res.path, width: res.width ?? 1, height: res.height ?? 1}], 'handwritten note');
      setPad(null);
      setStatus('Note added in the right margin. Save to put it in the Word file.');
    } catch (error) {
      setStatus(`The note could not be added: ${errorText(error)}`);
      log(`ink save failed: ${errorText(error)}`);
    }
  };

  const deleteNote = (id: string) => {
    const b = blocks.find(x => x.type === 'p' && x.runs.some(r => r.obj === 'ink' && r.ink === id)) as ParagraphBlock | undefined;
    setNoteOpen(null);
    setMenu(null);
    if (!b) {
      return;
    }
    commit([{op: 'inkDelete', para: b.index, id}], 'delete note');
    setStatus('Note deleted. Undo brings it back.');
  };

  // ---------------------------------------------------------------- citations (Zotero)

  const openCite = () => {
    // Where the citation goes: the caret stays after anything just typed.
    const at = caretAt;
    flushTyping();
    setCaret(at);
    setCite(c => ({...c, setup: !zotero, userId: zotero?.userId ?? '', apiKey: '', results: null, chosen: null, page: '', busy: false, message: ''}));
    setMenu('cite');
  };

  const saveZotero = async () => {
    const account = {userId: cite.userId.trim(), apiKey: cite.apiKey.trim() || zotero?.apiKey || ''};
    if (!/^\d+$/.test(account.userId)) {
      setCite(c => ({...c, message: 'The user ID is a number (on zotero.org → Settings → Security, above your keys) — not your username.'}));
      return;
    }
    if (!account.apiKey) {
      setCite(c => ({...c, message: 'Enter the API key.'}));
      return;
    }
    // Without it the Supernote blocks the request before it leaves (CT: "Could not reach Zotero").
    if (!(await ensureInternetPermission())) {
      setCite(c => ({...c, message: 'SNScribe needs permission to use the internet to reach Zotero. Allow it when asked.'}));
      return;
    }
    setCite(c => ({...c, busy: true, message: 'Checking with Zotero…'}));
    log(`zotero: checking user ${account.userId}`);
    try {
      await testZotero(account);
      await Docx?.store('zotero', JSON.stringify(account));
      setZotero(account);
      setCite(c => ({...c, setup: false, apiKey: '', busy: false, message: 'Connected to your Zotero library. Search for a source below.'}));
      setStatus('Connected to your Zotero library.');
    } catch (error) {
      setCite(c => ({...c, busy: false, message: errorText(error)}));
      setStatus(errorText(error));
    }
  };

  /** Load from file…: the key (and user ID, if there) from a .txt on the device — no typing. */
  const zoteroFromFile = async () => {
    try {
      if (!(await ensureFileReadPermission())) {
        setCite(c => ({...c, message: 'File access was not allowed.'}));
        return;
      }
      const picked = (await RattaFileSelector.selectFile({
        selectType: 0,
        maxNum: 1,
        title: 'Choose the text file with your Zotero key',
        rightButtonText: 'Use',
        suffixList: ['txt'],
      })) as string[] | null | undefined;
      log(`zotero key file: picker returned ${JSON.stringify(picked)}`);
      const path = picked?.find(x => typeof x === 'string' && x.length > 0);
      if (!path) {
        setCite(c => ({...c, message: 'No file was chosen.'}));
        return;
      }
      const text = await Docx!.readText(path);
      const found = credentialsIn(text);
      log(`zotero key file: ${text.length} chars, key ${found.apiKey ? 'found' : 'not found'}, user ID ${found.userId ? 'found' : 'not found'}`);
      if (!found.apiKey) {
        setCite(c => ({...c, message: 'No Zotero key in that file (a key is 24 letters and numbers).'}));
        return;
      }
      setCite(c => ({
        ...c,
        apiKey: found.apiKey!,
        userId: found.userId ?? c.userId,
        message: `Key loaded${found.userId ? ` and user ID ${found.userId}` : ''}. Tap Save. (You can delete the text file afterwards.)`,
      }));
    } catch (error) {
      setCite(c => ({...c, message: `Could not read the file: ${errorText(error)}`}));
    }
  };

  const searchCite = async (style = citeStyle) => {
    if (!zotero || !cite.query.trim()) {
      return;
    }
    if (!(await ensureInternetPermission())) {
      setCite(c => ({...c, message: 'SNScribe needs permission to use the internet to reach Zotero. Allow it when asked.'}));
      return;
    }
    setCite(c => ({...c, busy: true, chosen: null, message: 'Searching…'}));
    try {
      const results = await searchZotero(zotero, cite.query, CITE_STYLES.find(x => x.id === style)!.csl);
      log(`zotero search: ${results.length} result(s)${results.length ? `: ${results.slice(0, 3).map(r => `${r.authors} ${r.year}`).join('; ')}` : ''}`);
      const message = results.length ? `${results.length} found — tap one.` : `Nothing in your library matches “${cite.query.trim()}”.`;
      setCite(c => ({...c, results, busy: false, message}));
    } catch (error) {
      log(`zotero search failed: ${errorText(error)}`);
      setCite(c => ({...c, busy: false, message: errorText(error)}));
    }
  };

  /** Insert: the in-text citation at the caret, and the entry in the reference list — one undo step. */
  const insertCitation = () => {
    const src = cite.chosen;
    const at = caretAt;
    if (!src) {
      return;
    }
    if (!at) {
      setStatus('Tap where the citation goes first.');
      return;
    }
    const p = paragraph(at.para);
    const problem = p ? textEditProblem(p, at.offset, at.offset) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    const text = textOf(at.para);
    let cited = inText(src, citeStyle, cite.narrative, cite.page);
    if (at.offset > 0 && !/[\s(\[]/.test(text[at.offset - 1])) {
      cited = ` ${cited}`;
    }
    if (at.offset < text.length && /[A-Za-z0-9\u00C0-\u024F]/.test(text[at.offset])) {
      cited = `${cited} `;
    }
    const insert: Op = {op: 'text', para: at.para, start: at.offset, end: at.offset, text: cited};
    const {ops: refOps, added} = referenceOps(applyOps(blocks, [insert]), htmlToPieces(src.bibHtml), citeStyle);
    commit([insert, ...refOps], 'citation');
    setCaret({para: at.para, offset: at.offset + cited.length});
    setMenu(null);
    setStatus(added ? `Cited, and added to the reference list.` : 'Cited (already in the reference list).');
  };

  // ---------------------------------------------------------------- spelling

  /** Asks the dictionary about words it has not been asked about yet; marks follow. */
  const checkWords = async (words: string[]) => {
    const fresh = [...new Set(words.map(normal))].filter(w => !spellCache.current.has(w));
    if (fresh.length === 0 || !Docx) {
      return;
    }
    try {
      const bad = new Set(await Docx.spellCheck(SPELL_LANG, fresh));
      for (const w of fresh) {
        spellCache.current.set(w, !bad.has(w));
      }
      setSpellTick(t => t + 1);
    } catch (error) {
      log(`spell check failed: ${errorText(error)}`);
    }
  };

  // The page's words are checked shortly after the page settles or changes (typing included).
  useEffect(() => {
    if (!spellOn || !doc) {
      return;
    }
    const timer = setTimeout(() => {
      checkWords(window.flatMap(b => (b.type === 'p' ? tokens(paragraphText(b)).map(t => t.word) : [])));
    }, 500);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [window, spellOn, doc]);

  /** What to mark in a paragraph on the page (nothing when marking is off). */
  const spellFor = (b: (typeof window)[number]) => {
    if (!spellOn || b.type !== 'p') {
      return undefined;
    }
    const typingAt = typing && caretAt && caretAt.para === b.index ? caretAt.offset : undefined;
    const r = misspelledRanges(b, w => spellCache.current.get(w) === false, myWords, ignored, typingAt);
    return r.length ? r : undefined;
  };

  const spellItems = () =>
    blocks.flatMap(b =>
      b.type === 'p'
        ? misspelledRanges(b, w => spellCache.current.get(w) === false, myWords, ignored).map(r => ({
            para: b.index,
            ...r,
            word: paragraphText(b).slice(r.start, r.end),
          }))
        : [],
    );

  /** Shows misspelling `i`: selects it (the page turns to it) and fetches suggestions. */
  const showSpell = async (items: typeof sp.items, i: number) => {
    if (items.length === 0) {
      setSp(x => ({...x, items, at: 0, suggestions: [], replace: '', message: 'No spelling mistakes found.'}));
      return;
    }
    const k = Math.max(0, Math.min(i, items.length - 1));
    const m = items[k];
    flushTyping();
    showPlace(m.para, m.start);
    setCaret(null);
    setSelection({from: {para: m.para, offset: m.start}, to: {para: m.para, offset: m.end}});
    setSp(x => ({...x, items, at: k, suggestions: [], replace: '', message: `${k + 1} of ${items.length}`}));
    try {
      const suggestions = (await Docx?.spellSuggest(SPELL_LANG, normal(m.word))) ?? [];
      setSp(x => (x.items === items && x.at === k ? {...x, suggestions, replace: suggestions[0] ?? ''} : x));
    } catch (error) {
      log(`spell suggest failed: ${errorText(error)}`);
    }
  };

  /** Edit ▸ Spelling…: every misspelling in the document, one at a time. */
  const openSpelling = async () => {
    setMenu('spell');
    setSp(x => ({...x, items: [], message: 'Checking the whole document…'}));
    // The list is built after the dictionary has answered for every word (the answers are
    // kept in a ref, so they are there at once for spellItems).
    await checkWords(blocks.flatMap(b => (b.type === 'p' ? tokens(paragraphText(b)).map(t => t.word) : [])));
    showSpell(spellItems(), 0);
  };

  const spellChange = (all: boolean) => {
    const m = sp.items[sp.at];
    const to = sp.replace.trim();
    if (!m || !to) {
      return;
    }
    const targets = all ? sp.items.filter(x => x.word === m.word) : [m];
    const {ops, skipped} = replaceAllOps(blocks, targets, to);
    if (ops.length === 0) {
      setSp(x => ({...x, message: 'That word is inside a field; it can’t be changed here.'}));
      return;
    }
    commit(ops, all ? `spelling: all “${m.word}”` : 'spelling');
    setSelection(null);
    // The rest, as the document will be: later offsets in the same paragraph shift.
    const rest = sp.items
      .filter(x => !targets.includes(x))
      .map(x => {
        const before = targets.filter(t => t.para === x.para && t.end <= x.start).length;
        const delta = before * (to.length - m.word.length);
        return delta ? {...x, start: x.start + delta, end: x.end + delta} : x;
      });
    const next = rest.findIndex(x => x.para > m.para || (x.para === m.para && x.start > m.start));
    showSpell(rest, next < 0 ? 0 : next);
    if (skipped) {
      setSp(x => ({...x, message: `${x.message} (${skipped} inside fields left alone)`}));
    }
  };

  const spellSkip = (word?: string) => {
    const m = sp.items[sp.at];
    if (!m) {
      return;
    }
    const rest = word ? sp.items.filter(x => x.word !== word) : sp.items.filter(x => x !== m);
    const next = rest.findIndex(x => x.para > m.para || (x.para === m.para && x.start > m.start));
    showSpell(rest, next < 0 ? 0 : next);
  };

  const spellIgnoreAll = () => {
    const m = sp.items[sp.at];
    if (m) {
      setIgnored(s0 => new Set([...s0, normal(m.word)]));
      spellSkip(m.word);
    }
  };

  const spellAdd = () => {
    const m = sp.items[sp.at];
    if (!m) {
      return;
    }
    const next = new Set([...myWords, normal(m.word).toLowerCase()]);
    setMyWords(next);
    Docx?.store(`spell-words-${SPELL_LANG}`, JSON.stringify([...next]));
    spellSkip(m.word);
  };

  // ---------------------------------------------------------------- quotes from PDFs / EPUBs

  const openQuotes = async () => {
    const at = caretAt;
    flushTyping();
    setCaret(at);
    const quotes = parseList<SavedQuote>(await Docx?.load(QUOTES_KEY));
    const sources = parseMap<SourceRef>(await Docx?.load(SOURCES_KEY));
    setQp(q => ({...q, mode: 'list', quotes, sources, chosen: null, source: null, results: [], message: quotes.length ? '' : 'No quotes yet. In a PDF or EPUB, select text and tap Quote → SNScribe.'}));
    setMenu('quotes');
  };

  const saveQuotes = (quotes: SavedQuote[]) => {
    Docx?.store(QUOTES_KEY, JSON.stringify(quotes));
    setQp(q => ({...q, quotes}));
  };

  /** A source's citation in the current style: from Zotero, or formatted from its details. */
  const resolveSource = async (ref: SourceRef): Promise<Source> => {
    if (ref.kind === 'details') {
      return formatSource(ref.details, citeStyle);
    }
    if (!zotero) {
      throw new Error('Connect Zotero first (Edit → Cite from Zotero…).');
    }
    if (!(await ensureInternetPermission())) {
      throw new Error('SNScribe needs permission to use the internet to reach Zotero.');
    }
    return zoteroItem(zotero, ref.key, CITE_STYLES.find(x => x.id === citeStyle)!.csl);
  };

  const chooseQuote = async (quote: SavedQuote) => {
    const ref = qp.sources[quote.path];
    setQp(q => ({...q, chosen: quote, page: String(quote.page), doi: quote.doi ?? '', query: fileName(quote.path).replace(/\.(pdf|epub)$/i, '').replace(/[_]+/g, ' '), results: [], message: ''}));
    if (!ref) {
      setQp(q => ({...q, mode: 'source'}));
      return;
    }
    setQp(q => ({...q, busy: true, message: 'Getting the citation…'}));
    try {
      const source = await resolveSource(ref);
      setQp(q => ({...q, mode: 'preview', source, busy: false, message: ''}));
    } catch (error) {
      setQp(q => ({...q, mode: 'source', busy: false, message: errorText(error)}));
    }
  };

  /** Remember where this file's details come from, and go on to the preview. */
  const takeSource = async (ref: SourceRef) => {
    const quote = qp.chosen;
    if (!quote) {
      return;
    }
    const sources = {...qp.sources, [quote.path]: ref};
    Docx?.store(SOURCES_KEY, JSON.stringify(sources));
    setQp(q => ({...q, sources, busy: true, message: 'Getting the citation…'}));
    try {
      const source = await resolveSource(ref);
      setQp(q => ({...q, mode: 'preview', source, busy: false, message: ''}));
    } catch (error) {
      setQp(q => ({...q, busy: false, message: errorText(error)}));
    }
  };

  const sourceFromDoi = async () => {
    const doi = qp.doi.trim();
    if (!doi) {
      return;
    }
    if (!(await ensureInternetPermission())) {
      setQp(q => ({...q, message: 'SNScribe needs permission to use the internet to look up the DOI.'}));
      return;
    }
    setQp(q => ({...q, busy: true, message: `Looking up ${doi}…`}));
    try {
      const details = await lookUpDoi(doi, line => log(line));
      await takeSource({kind: 'details', details});
    } catch (error) {
      setQp(q => ({...q, busy: false, message: errorText(error)}));
    }
  };

  const sourceFromEpub = async () => {
    const quote = qp.chosen;
    if (!quote || !Docx) {
      return;
    }
    try {
      const details = detailsFromEpub(await Docx.epubInfo(quote.path));
      if (!details.title) {
        setQp(q => ({...q, message: 'This EPUB has no title in its details. Type them instead.'}));
        return;
      }
      // Shown in the form first: EPUB details are sometimes thin (no year, "Unknown" author).
      setQp(q => ({
        ...q,
        mode: 'form',
        form: {
          ...q.form,
          type: 'book',
          authors: details.authors.map(a => (a.given ? `${a.family}, ${a.given}` : a.family)).join('\n'),
          year: details.year ?? '',
          title: details.title,
          publisher: details.publisher ?? '',
        },
        message: 'From the EPUB — check and complete, then Use these details.',
      }));
    } catch (error) {
      setQp(q => ({...q, message: `Could not read the EPUB: ${errorText(error)}`}));
    }
  };

  const searchQuoteZotero = async () => {
    if (!zotero) {
      setQp(q => ({...q, message: 'Connect Zotero first (Edit → Cite from Zotero…).'}));
      return;
    }
    if (!(await ensureInternetPermission())) {
      setQp(q => ({...q, message: 'SNScribe needs permission to use the internet to reach Zotero.'}));
      return;
    }
    setQp(q => ({...q, busy: true, message: 'Searching…'}));
    try {
      const results = await searchZotero(zotero, qp.query, CITE_STYLES.find(x => x.id === citeStyle)!.csl);
      setQp(q => ({...q, results, busy: false, message: results.length ? `${results.length} found — tap the right one.` : 'Nothing matches. Try the author or a title word.'}));
    } catch (error) {
      setQp(q => ({...q, busy: false, message: errorText(error)}));
    }
  };

  const formDetails = (): Details | null => {
    const f = qp.form;
    if (!f.title.trim()) {
      setQp(q => ({...q, message: 'A title is needed.'}));
      return null;
    }
    const link = f.link.trim();
    return {
      type: f.type,
      authors: f.authors
        .split('\n')
        .map(l => l.trim())
        .filter(Boolean)
        .map(l => {
          const [family, given] = l.split(',', 2).map(x => x.trim());
          return given ? {family, given} : {family};
        }),
      year: f.year.trim() || undefined,
      title: f.title.trim(),
      container: f.container.trim() || undefined,
      volume: f.volume.trim() || undefined,
      issue: f.issue.trim() || undefined,
      pages: f.pages.trim() || undefined,
      publisher: f.publisher.trim() || undefined,
      ...(/^10\./.test(link) || /doi\.org\//i.test(link) ? {doi: link} : link ? {url: link} : {}),
    };
  };

  /** Insert: the quote with its citation at the caret (or as a block quote), and the reference. */
  const insertQuote = () => {
    const quote = qp.chosen;
    const source = qp.source;
    const at = caretAt;
    if (!quote || !source) {
      return;
    }
    if (!at) {
      setQp(q => ({...q, message: 'Tap where the quote goes first (close this, tap the page, open Insert quote again).'}));
      return;
    }
    const p = paragraph(at.para);
    const problem = p ? textEditProblem(p, at.offset, at.offset) : 'Paragraph not found.';
    if (problem) {
      setQp(q => ({...q, message: problem}));
      return;
    }
    const r = quoteWithReference(blocks, applyOps, at, quote.text, source, citeStyle, qp.page);
    commit(r.ops, 'quote');
    setCaret(r.caret);
    setMenu(null);
    setStatus(`${r.block ? 'Block quote' : 'Quote'} inserted${r.added ? ', and the source added to the reference list' : ''}.`);
  };

  // ---------------------------------------------------------------- comments

  const initialsOf = (name: string) =>
    name
      .split(/\s+/)
      .filter(Boolean)
      .map(w => w[0].toUpperCase())
      .join('')
      .slice(0, 3);

  /** Comment…: on the selection, or at the caret. */
  const startComment = () => {
    const ranges = !typing && selection ? rangesBetween(blocks, selection.from, selection.to) : [];
    let form: {fromPara: number; from: number; toPara: number; to: number} | null = null;
    if (ranges.length > 0) {
      const a = ranges[0];
      const z = ranges[ranges.length - 1];
      form = {fromPara: a.para, from: a.start, toPara: z.para, to: z.end};
    } else if (caretAt) {
      flushTyping();
      form = {fromPara: caretAt.para, from: caretAt.offset, toPara: caretAt.para, to: caretAt.offset};
    }
    if (!form) {
      setMenu(null);
      setStatus('Select the words to comment on, or tap where the comment goes.');
      return;
    }
    const quote = ranges.map(r => textOf(r.para).slice(r.start, r.end)).join(' ');
    setCommentForm({...form, text: '', quote});
    setMenu('comment');
  };

  const addComment = () => {
    if (!commentForm) {
      return;
    }
    const text = commentForm.text.trim();
    if (!text) {
      setStatus('Write the comment first.');
      return;
    }
    const {quote: _q, text: _t, ...at} = commentForm;
    const name = author.trim();
    commit(
      [{op: 'comment', para: -1, id: nextCommentId(comments), ...at, text, author: name, initials: initialsOf(name), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z')}],
      'comment',
    );
    setCommentForm(null);
    setMenu(null);
    setStatus('Comment added. It is saved with the document as a Word comment.');
  };

  const openThread = (id: string) => {
    setThread(id);
    setReplyText('');
    setMenu('thread');
  };

  const reply = () => {
    const text = replyText.trim();
    if (!thread || !text) {
      return;
    }
    const name = author.trim();
    commit(
      [
        {
          op: 'comment',
          para: -1,
          id: nextCommentId(comments),
          fromPara: 0,
          from: 0,
          toPara: 0,
          to: 0,
          text,
          author: name,
          initials: initialsOf(name),
          date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'),
          parent: Number(thread),
        },
      ],
      'reply',
    );
    setReplyText('');
  };

  const deleteThread = () => {
    if (!thread) {
      return;
    }
    commit([{op: 'uncomment', para: -1, ids: threadIds(comments, thread)}], 'delete comment');
    setThread(null);
    setMenu(null);
    setStatus('Comment deleted. Undo brings it back.');
  };

  // ---------------------------------------------------------------- tracked changes

  /**
   * Accept or reject one tracked change (or all: para -1, id '*'). The file decides the
   * result (DocxModule.preview), so the screen shows exactly what Save will write.
   */
  const review = async (para: number, id: string, accept: boolean) => {
    if (!doc?.source || !Docx) {
      return;
    }
    flushTyping();
    const all = para < 0;
    const paras = all ? blocks.filter((b): b is ParagraphBlock => b.type === 'p' && !!b.revs?.length).map(b => b.index) : [para];
    if (paras.length === 0) {
      setStatus('There are no tracked changes to review.');
      return;
    }
    const op = {op: 'revision' as const, para, id, accept};
    setBusy(true);
    try {
      const res = await Docx.preview(doc.source, [...applied, op], paras);
      const result: Record<number, {runs: ParagraphBlock['runs']; revs?: ParagraphBlock['revs']}> = {};
      for (const b of res.blocks) {
        result[b.index] = {runs: b.runs, revs: b.revs};
      }
      commit([{...op, result}], `${accept ? 'accept' : 'reject'} ${all ? 'all changes' : 'change'}`);
      setStatus(all ? `${accept ? 'Accepted' : 'Rejected'} every tracked change in the text. Undo puts them back.` : accept ? 'Change accepted.' : 'Change rejected.');
    } catch (error) {
      setStatus(`Could not review the change: ${errorText(error)}`);
      log(`review failed: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  // ---------------------------------------------------------------- header, page numbers, links

  const headerNow = () => headerFooterAfter(doc?.header, 'header', applied);
  const footerNow = () => headerFooterAfter(doc?.footer, 'footer', applied);

  const openHeaderForm = () => {
    const line = (hf: ReturnType<typeof headerNow>, align: HfLine['align']): HfLine =>
      hf ? {text: hf.text, align: hf.align === 'justify' ? 'left' : hf.align, page: hf.pageNumber} : {text: '', align, page: false};
    const header = line(headerNow(), 'right');
    const footer = line(footerNow(), 'center');
    setHfForm({header, footer, was: {header, footer}});
    setMenu('header');
  };

  /** Writes the header and/or footer — only the ones that were changed in the form. */
  const applyHeader = () => {
    if (!hfForm) {
      return;
    }
    const ops: Op[] = [];
    for (const kind of ['header', 'footer'] as const) {
      const now = hfForm[kind];
      const was = hfForm.was[kind];
      if (now.text.trim() !== was.text.trim() || now.align !== was.align || now.page !== was.page) {
        ops.push({op: 'headerFooter', para: -1, kind, text: now.text.trim(), pageNumber: now.page, align: now.align});
      }
    }
    setMenu(null);
    setHfForm(null);
    if (ops.length === 0) {
      setStatus('Header and footer unchanged.');
      return;
    }
    commit(ops, 'header/footer');
    setStatus('Saved with the document; Word shows it on every page. (The Supernote page does not draw headers or footers.)');
  };

  /** Link…: the selection (within one paragraph) becomes a link; its text is offered as the address. */
  const startLink = () => {
    const ranges = !typing && selection ? rangesBetween(blocks, selection.from, selection.to).filter(r => r.end > r.start) : [];
    if (ranges.length !== 1) {
      setMenu(null);
      setStatus(ranges.length > 1 ? 'A link has to stay within one paragraph.' : 'Select the words to link first.');
      return;
    }
    const r = ranges[0];
    const p = paragraph(r.para);
    const problem = p ? linkProblem(p, r.start, r.end) : 'Paragraph not found.';
    if (problem) {
      setMenu(null);
      setStatus(problem);
      return;
    }
    const words = textOf(r.para).slice(r.start, r.end);
    setLinkForm({...r, url: linkUrl(words) ? words.trim() : ''});
    setMenu('link');
  };

  const applyLink = () => {
    if (!linkForm) {
      return;
    }
    const url = linkUrl(linkForm.url);
    if (!url) {
      setStatus('That isn\'t a web address or DOI (like https://… or 10.1037/…).');
      return;
    }
    setMenu(null);
    setLinkForm(null);
    commit([{op: 'link', para: linkForm.para, start: linkForm.start, end: linkForm.end, url}], 'link');
    setStatus(`Linked to ${url}`);
  };

  /** Remove link: the link at the caret, or the links the selection touches. */
  const removeLink = () => {
    setMenu(null);
    const ranges: Range[] =
      !typing && selection ? rangesBetween(blocks, selection.from, selection.to) : caretAt ? [{para: caretAt.para, start: caretAt.offset, end: caretAt.offset}] : [];
    flushTyping();
    const ops: Op[] = [];
    for (const r of ranges) {
      const p = paragraph(r.para);
      if (p && linkSpan(p, r.start, r.end)) {
        ops.push({op: 'unlink', para: r.para, start: r.start, end: r.end});
      }
    }
    if (ops.length === 0) {
      setStatus('There is no link here.');
      return;
    }
    commit(ops, 'remove link');
  };

  /** Paragraph formatting for the targeted paragraphs, as one undo step. The menu stays open. */
  const paraTool = (props: Omit<Extract<Op, {op: 'para'}>, 'op' | 'para'>, label: string) => {
    const paras = targetParagraphs();
    if (paras.length === 0) {
      setStatus('Tap in a paragraph first.');
      return;
    }
    commit(paras.map(para => ({op: 'para', para, ...props})), label);
  };

  /** Insert page break: the paragraph splits at the caret and its second half starts a new page. */
  const insertPageBreak = () => {
    const at = caretAt;
    flushTyping();
    if (!at) {
      setStatus('Tap where the new page should start.');
      return;
    }
    if (at.offset === 0) {
      commit([{op: 'para', para: at.para, pb: true}], 'page break');
      return;
    }
    const p = paragraph(at.para);
    const problem = p ? splitProblem(p, at.offset) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    commit(
      [
        {op: 'split', para: at.para, offset: at.offset},
        {op: 'para', para: at.para + 1, pb: true},
      ],
      'page break',
    );
    setCaret({para: at.para + 1, offset: 0});
  };

  /** The paragraph the Paragraph menu shows the settings of. */
  const paraShown = (): ParagraphBlock | undefined => {
    const para = !typing && selection ? Math.min(selection.from.para, selection.to.para) : caretAt?.para;
    return para === undefined ? undefined : paragraph(para);
  };



  const format = (prop: FormatProp, label: string) => {
    const ranges = targets();
    if (ranges.length === 0) {
      nothingSelected();
      return;
    }
    commit(formatOps(blocks, ranges, prop), label);
  };

  /** HL ▾: highlight the selection (or the word at the caret) in a colour, or remove highlighting. */
  const highlight = (colour: string | null, label: string) => {
    const ranges = targets();
    setMenu(null);
    if (ranges.length === 0) {
      nothingSelected();
      return;
    }
    commit(highlightOps(ranges, colour), label);
  };

  // ---------------------------------------------------------------- find & replace

  const findAll = () => findMatches(blocks, find.query, find.matchCase);

  /** Shows and selects match `i` (the page turns to it when it is elsewhere). */
  const showMatch = (i: number) => {
    const matches = findAll();
    if (matches.length === 0) {
      setFind(f => ({...f, at: -1}));
      setStatus(find.query ? `“${find.query}” was not found.` : 'Type what to find.');
      return;
    }
    const k = ((i % matches.length) + matches.length) % matches.length;
    const m = matches[k];
    flushTyping();
    showPlace(m.para, m.start);
    setCaret(null);
    setSelection({from: {para: m.para, offset: m.start}, to: {para: m.para, offset: m.end}});
    setFind(f => ({...f, at: k}));
    setStatus(`${k + 1} of ${matches.length}`);
  };

  /** Replace: the shown match (if it is still selected), then on to the next. */
  const replaceOne = () => {
    const matches = findAll();
    const m = find.at >= 0 ? matches[find.at] : undefined;
    const sel = selection;
    if (!m || !sel || sel.from.para !== m.para || Math.min(sel.from.offset, sel.to.offset) !== m.start) {
      showMatch(find.at + 1);
      return;
    }
    const p = paragraph(m.para);
    const problem = p ? textEditProblem(p, m.start, m.end) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      showMatch(find.at + 1);
      return;
    }
    commit([{op: 'text', para: m.para, start: m.start, end: m.end, text: find.replace}], 'replace');
    setSelection(null);
    setStatus(`Replaced 1. ${matches.length - 1} left — Next to go on.`);
    setFind(f => ({...f, at: f.at - 1}));
  };

  const replaceAll = () => {
    const matches = findAll();
    if (matches.length === 0) {
      setStatus(`“${find.query}” was not found.`);
      return;
    }
    const {ops, skipped} = replaceAllOps(blocks, matches, find.replace);
    setSelection(null);
    commit(ops, `replace all “${find.query}”`);
    setFind(f => ({...f, at: -1}));
    setStatus(`Replaced ${ops.length}${skipped ? `; ${skipped} inside fields left alone` : ''}. Undo puts them all back.`);
  };

  /**
   * Superscript / Subscript: on a selection, formats it. With only a caret, switches it on
   * or off for what is typed next — nothing already written changes.
   */
  const scriptTool = (prop: 'sup' | 'sub') => {
    setMenu(null);
    if (!typing && selection) {
      format(prop, prop === 'sup' ? 'superscript' : 'subscript');
      return;
    }
    if (!caretAt) {
      setStatus('Tap where to type, or select the text first.');
      return;
    }
    const now = typingScript();
    // Keep typing where it left off: saving what was typed must leave the caret after it
    // (it stayed where typing began, so the next character went there instead).
    const at = caretAt;
    flushTyping();
    setCaret(at);
    const next = now === prop ? 'none' : prop;
    setScript(next);
    setStatus(next === 'none' ? 'Back to normal text for what you type next.' : `${prop === 'sup' ? 'Superscript' : 'Subscript'} on for what you type next — choose it again to turn it off.`);
  };

  /** What typing at the caret produces now: the switch, else the text just before the caret. */
  const typingScript = (): 'sup' | 'sub' | 'none' => {
    if (script) {
      return script;
    }
    const at = caretAt;
    const p = at ? paragraph(at.para) : undefined;
    if (!p || !at) {
      return 'none';
    }
    let offset = 0;
    for (const r of p.runs) {
      const end = offset + r.t.length;
      if (at.offset > offset && at.offset <= end) {
        return r.sup ? 'sup' : r.sub ? 'sub' : 'none';
      }
      offset = end;
    }
    return 'none';
  };

  const style = (kind: StyleKind, label: string) => {
    const paras = targetParagraphs();
    if (paras.length === 0) {
      nothingSelected();
      return;
    }
    commit(styleOps(paras.map(para => ({para, start: 0, end: 0})), kind, looksAfter(doc?.looks, applied)[kind]), label);
  };

  /** Numbered / bulleted / not a list, for the paragraphs targeted. A list just above is continued. */
  const listTool = (kind: 'number' | 'bullet' | 'none') => {
    const paras = targetParagraphs();
    if (paras.length === 0) {
      nothingSelected();
      return;
    }
    let listId = '';
    if (kind !== 'none') {
      const first = blocks.findIndex(b => b.type === 'p' && b.index === paras[0]);
      const prev = blocks[first - 1];
      listId = prev?.type === 'p' && prev.num && listKind(prev.num.id) === kind ? String(prev.num.id) : `${kind[0]}${++listCounter.current}`;
    }
    commit(paras.map(para => ({op: 'list', para, kind, listId})), kind === 'none' ? 'not a list' : `${kind} list`);
  };

  const paragraph = (para: number) => blocks.find(b => b.type === 'p' && b.index === para) as ParagraphBlock | undefined;

  /** Deletes the selection (joining paragraphs it spans). Returns where the caret is left. */
  const remove = (): Pos | null => {
    if (!selection) {
      return null;
    }
    const ranges = rangesBetween(blocks, selection.from, selection.to);
    const ops: Op[] = [];
    for (const r of ranges) {
      const p = paragraph(r.para);
      const problem = p && textEditProblem(p, r.start, r.end);
      if (!p || problem) {
        setStatus(problem ?? 'Paragraph not found.');
        return null;
      }
      // Within one paragraph, also tidy the space a deleted word leaves.
      const d = ranges.length === 1 ? deletionRange(paragraphText(p), r.start, r.end) : r;
      ops.push({op: 'text', para: r.para, start: d.start, end: d.end, text: ''});
    }
    // Across paragraphs, what is left joins into one, as in Word: last onto first.
    const first = selection.from.para < selection.to.para ? selection.from.para : selection.to.para;
    const last = selection.from.para < selection.to.para ? selection.to.para : selection.from.para;
    for (let para = last; para > first; para--) {
      const problem = joinProblem(blocks, para);
      if (problem) {
        setStatus(`${problem} Delete within the text on one side of it.`);
        return null;
      }
      ops.push({op: 'join', para});
    }
    if (ops.length === 0) {
      return null;
    }
    commit(ops, 'delete');
    setSelection(null);
    const first0 = ops[0];
    const at = {para: first, offset: first0.op === 'text' && first0.para === first ? first0.start : paragraphText(paragraph(first)!).length};
    setCaret(at);
    return at;
  };

  const startReplace = () => {
    if (!selection) {
      return;
    }
    const ranges = rangesBetween(blocks, selection.from, selection.to);
    if (ranges.length !== 1) {
      setStatus('Replace works inside one paragraph. Select less, or use Delete.');
      return;
    }
    const r = ranges[0];
    const p = paragraph(r.para);
    const problem = p ? textEditProblem(p, r.start, r.end) : 'Paragraph not found.';
    if (problem) {
      setStatus(problem);
      return;
    }
    setInput('');
    setTyping({mode: 'replace', range: r});
    keys.current?.focus();
  };

  /** Commits what was typed (one undo step) and ends typing. */
  const flushTyping = () => {
    if (!typing) {
      return;
    }
    setTyping(null);
    setInput('');
    if (pending.length > 0) {
      commit(pending, typing.mode);
      setSelection(null);
    }
  };

  /** Enter: commit what was typed, start a new paragraph at the caret, keep typing there. */
  const enter = () => {
    if (!typing || !caretAt) {
      return;
    }
    const at = caretAt;
    const p = paragraph(at.para);
    if (pending.length === 0 && p?.num && paragraphText(p) === '' && endListAt(at)) {
      setTyping({mode: 'insert', at});
      return;
    }
    const problem = p ? splitProblem(p, at.offset) : 'Paragraph not found.';
    flushTyping();
    setSelection(null);
    if (problem) {
      setStatus(problem);
      setTyping({mode: 'insert', at});
      return;
    }
    commit([{op: 'split', para: at.para, offset: at.offset}], 'new paragraph');
    setTyping({mode: 'insert', at: {para: at.para + 1, offset: 0}});
  };

  /** Hide keyboard (Edit menu): commit, leave the caret after the new text, and put the keyboard away. */
  const done = () => {
    const at = caretAt;
    flushTyping();
    setCaret(at);
    keys.current?.blur();
    Keyboard.dismiss();
  };

  /**
   * Backspace the text field can't handle itself (nothing typed in it): deletes the
   * selection, or the character before the caret, or joins onto the previous paragraph.
   */
  const backspace = () => {
    if (input !== '') {
      return;
    }
    if (!typing) {
      if (selection) {
        remove();
      } else if (caret) {
        setTyping({mode: 'insert', at: caret});
        backspaceAt(caret);
      }
      return;
    }
    if (typing.mode === 'replace') {
      const r = typing.range;
      commit([{op: 'text', para: r.para, start: r.start, end: r.end, text: ''}], 'delete');
      setSelection(null);
      setTyping({mode: 'insert', at: {para: r.para, offset: r.start}});
      return;
    }
    backspaceAt(typing.at);
  };

  /** Backspace at a caret while typing: the character before it, or a paragraph join. */
  /**
   * The edit that removes the page break just before `at`, or null: the paragraph's own
   * "starts on a new page", or a Word page break (its line break) right before it — in the
   * same paragraph, or ending the paragraph before.
   */
  const pageBreakRemoval = (at: Pos): Op[] | null => {
    const p = paragraph(at.para);
    if (!p) {
      return null;
    }
    if (at.offset === 0 && p.pb) {
      return [{op: 'para', para: at.para, pb: false}];
    }
    if (at.offset > 0) {
      return pageBreakChars(p).includes(at.offset - 1) ? [{op: 'text', para: at.para, start: at.offset - 1, end: at.offset, text: ''}] : null;
    }
    const prev = paragraph(at.para - 1);
    const len = prev ? paragraphText(prev).length : 0;
    return prev && len > 0 && pageBreakChars(prev).includes(len - 1) ? [{op: 'text', para: prev.index, start: len - 1, end: len, text: ''}] : null;
  };

  const backspaceAt = (at: Pos) => {
    const p = paragraph(at.para);
    if (!p) {
      return;
    }
    if (at.offset === 0 && p.num) {
      // At the start of a list item: the number or bullet goes first, as in Word.
      endListAt(at);
      setTyping({mode: 'insert', at});
      return;
    }
    // At the start of a page: the page break goes first, as in Word; a second Backspace joins.
    const unbreak = at.offset === 0 ? pageBreakRemoval(at) : null;
    if (unbreak) {
      commit(unbreak, 'remove page break');
      setTyping({mode: 'insert', at});
      setStatus('Page break removed. Backspace again joins this paragraph to the one before.');
      return;
    }
    if (at.offset === 0) {
      // At the start of a paragraph: join it onto the one before, caret at the join.
      const problem = joinProblem(blocks, at.para);
      if (problem) {
        setStatus(problem);
        return;
      }
      const prev = paragraph(at.para - 1)!;
      commit([{op: 'join', para: at.para}], 'join paragraphs');
      setTyping({mode: 'insert', at: {para: at.para - 1, offset: paragraphText(prev).length}});
      return;
    }
    const problem = textEditProblem(p, at.offset - 1, at.offset);
    if (problem) {
      setStatus(problem);
      return;
    }
    commit([{op: 'text', para: at.para, start: at.offset - 1, end: at.offset, text: ''}], 'backspace');
    setTyping({mode: 'insert', at: {para: at.para, offset: at.offset - 1}});
  };

  /** Every change of the invisible field: the first key decides what it does. */
  /** New lists made while editing get names "n1", "b2" … (the writer gives each its own list). */
  const listCounter = useRef(0);

  /**
   * Word's automatic lists: "1." / "1)" / "-" / "*" and a space typed at the start of a
   * body paragraph turn it into a numbered or bulleted list item. Two undo steps, like
   * Word: the typed marker, then the conversion — so Undo brings the marker back.
   * A list made just above is continued; otherwise a new one starts at 1.
   */
  const autoList = (text: string): boolean => {
    if (!typing || typing.mode !== 'insert' || typing.at.offset !== 0) {
      return false;
    }
    const m = /^(1[.)]|[-*]) $/.exec(text);
    const p = paragraph(typing.at.para);
    if (!m || !p || p.num || p.kind !== 'body') {
      return false;
    }
    const kind = m[1].startsWith('1') ? 'number' : 'bullet';
    const prev = blocks[blocks.indexOf(p) - 1];
    const listId =
      prev?.type === 'p' && prev.num && listKind(prev.num.id) === kind ? String(prev.num.id) : `${kind[0]}${++listCounter.current}`;
    const para = typing.at.para;
    commit([{op: 'text', para, start: 0, end: 0, text}], 'typing');
    commit(
      [
        {op: 'text', para, start: 0, end: text.length, text: ''},
        {op: 'list', para, kind, listId},
      ],
      'automatic list',
    );
    setInput('');
    setTyping({mode: 'insert', at: {para, offset: 0}});
    return true;
  };

  /** Enter or Backspace on an empty list item ends the list there, as in Word. */
  const endListAt = (at: Pos): boolean => {
    const p = paragraph(at.para);
    if (!p?.num || (paragraphText(p) !== '' && at.offset !== 0)) {
      return false;
    }
    commit([{op: 'list', para: at.para, kind: 'none', listId: ''}], 'end list');
    return true;
  };

  const onKeysText = (text: string) => {
    if (typing) {
      if (!autoList(text)) {
        setInput(text);
      }
      return;
    }
    if (text === '') {
      return;
    }
    if (selection) {
      const ranges = rangesBetween(blocks, selection.from, selection.to);
      const r = ranges[0];
      const p = r && paragraph(r.para);
      if (ranges.length === 1 && p && !textEditProblem(p, r.start, r.end)) {
        setTyping({mode: 'replace', range: r}); // type over the selection
        setInput(text);
        return;
      }
      const at = remove(); // across paragraphs: delete (and join), then type where it leaves off
      if (at) {
        setTyping({mode: 'insert', at});
        setInput(text);
      }
      return;
    }
    if (caret) {
      const p = paragraph(caret.para);
      const problem = p ? textEditProblem(p, caret.offset, caret.offset) : 'Paragraph not found.';
      if (problem) {
        setStatus(problem);
        return;
      }
      setTyping({mode: 'insert', at: caret});
      setInput(text);
    }
  };

  /** Enter: a new paragraph at the caret (replacing the selection, if any). */
  const onEnter = () => {
    if (typing) {
      enter();
      return;
    }
    let at = caret;
    if (!selection && at && paragraph(at.para)?.num && textOf(at.para) === '' && endListAt(at)) {
      return;
    }
    if (selection) {
      at = remove();
    } else if (at) {
      const p = paragraph(at.para);
      const problem = p ? splitProblem(p, at.offset) : 'Paragraph not found.';
      if (problem) {
        setStatus(problem);
        return;
      }
    }
    if (!at) {
      return;
    }
    commit([{op: 'split', para: at.para, offset: at.offset}], 'new paragraph');
    setCaret({para: at.para + 1, offset: 0});
  };

  // ---------------------------------------------------------------- keyboard

  /** Where the caret or the moving end of the selection is, committing any typing first. */
  const cursorForKeys = (): Pos | null => {
    const at = caretAt;
    flushTyping();
    return at;
  };

  /** Delete (forward): the selection, else the character after the caret, else joins the next paragraph on. */
  const forwardDelete = () => {
    if (!typing && selection) {
      remove();
      return;
    }
    const at = cursorForKeys();
    const p = at && paragraph(at.para);
    if (!at || !p) {
      return;
    }
    if (at.offset < paragraphText(p).length) {
      const problem = textEditProblem(p, at.offset, at.offset + 1);
      if (problem) {
        setStatus(problem);
        return;
      }
      commit([{op: 'text', para: at.para, start: at.offset, end: at.offset + 1, text: ''}], 'delete');
    } else {
      const problem = joinProblem(blocks, at.para + 1);
      if (problem) {
        setStatus(problem);
        return;
      }
      commit([{op: 'join', para: at.para + 1}], 'join paragraphs');
    }
    setCaret(at);
  };

  /** One step of caret movement. Up/Down ask the paragraph's own layout for the line above or below. */
  const step = async (pos: Pos, key: string): Promise<Pos | null> => {
    const len = textOf(pos.para).length;
    const prevEnd = paragraph(pos.para - 1) ? {para: pos.para - 1, offset: textOf(pos.para - 1).length} : null;
    const nextStart = paragraph(pos.para + 1) ? {para: pos.para + 1, offset: 0} : null;
    switch (key) {
      case 'LEFT':
        return pos.offset > 0 ? {para: pos.para, offset: pos.offset - 1} : prevEnd;
      case 'RIGHT':
        return pos.offset < len ? {para: pos.para, offset: pos.offset + 1} : nextStart;
      case 'HOME':
        return {para: pos.para, offset: 0};
      case 'END':
        return {para: pos.para, offset: len};
      default: {
        const up = key === 'UP';
        const i = window.findIndex(b => b.type === 'p' && b.index === pos.para);
        const tag = i >= 0 ? findNodeHandle(textRefs.current[i] ?? null) : null;
        if (tag !== null && DocxText) {
          const b = window[i] as ParagraphBlock;
          const r = await DocxText.lineMove(tag, toShown(b, pos.offset), up ? -1 : 1);
          if (r.offset !== undefined) {
            return {para: pos.para, offset: fromShown(b, r.offset)};
          }
        }
        return up ? prevEnd : nextStart;
      }
    }
  };

  /** Arrows, Home, End; with Shift they extend the selection from where it started. */
  const move = async (key: string, shift: boolean) => {
    setScript(null);
    const sel = typing ? null : selection;
    const from = cursorForKeys();
    const cur = sel ? sel.to : from;
    if (!cur) {
      return;
    }
    if (!shift && sel) {
      // Collapse the selection to the side the arrow points to.
      const [s, e] = comparePos(sel.from, sel.to) <= 0 ? [sel.from, sel.to] : [sel.to, sel.from];
      setSelection(null);
      setCaret(key === 'LEFT' || key === 'UP' || key === 'HOME' ? s : e);
      return;
    }
    const next = await step(cur, key);
    if (!next) {
      return;
    }
    if (!shift) {
      setSelection(null);
      setCaret(next);
      return;
    }
    const anchorPos = sel ? sel.from : cur;
    if (comparePos(anchorPos, next) === 0) {
      setSelection(null);
      setCaret(next);
    } else {
      setSelection({from: anchorPos, to: next});
      setCaret(null);
    }
  };

  const selectedText = (): string =>
    selection
      ? rangesBetween(blocks, selection.from, selection.to)
          .map(r => textOf(r.para).slice(r.start, r.end))
          .join('\n')
      : '';

  const copy = async (cut: boolean) => {
    if (typing || !selection) {
      nothingSelected();
      return;
    }
    const ok = await DocxKeys?.copy(selectedText());
    setStatus(ok ? (cut ? 'Cut.' : 'Copied.') : 'Could not use the clipboard.');
    if (ok && cut) {
      remove();
    }
  };

  /** Paste plain text at the caret (or over the selection); line breaks become new paragraphs. One undo step. */
  const paste = (raw: string) => {
    const lines = raw.replace(/\r\n?/g, '\n').split('\n').map(clean);
    let at: Pos | null;
    if (typing) {
      at = cursorForKeys();
    } else if (selection) {
      at = remove();
    } else {
      at = caret;
      const p = at && paragraph(at.para);
      const problem = p && at ? textEditProblem(p, at.offset, at.offset) : null;
      if (problem) {
        setStatus(problem);
        return;
      }
    }
    if (!at) {
      return;
    }
    const ops: Op[] = [];
    let pos = at;
    lines.forEach((line, i) => {
      if (line) {
        ops.push({op: 'text', para: pos.para, start: pos.offset, end: pos.offset, text: line});
        pos = {para: pos.para, offset: pos.offset + line.length};
      }
      if (i < lines.length - 1) {
        ops.push({op: 'split', para: pos.para, offset: pos.offset});
        pos = {para: pos.para + 1, offset: 0};
      }
    });
    commit(ops, 'paste');
    setSelection(null);
    setCaret(pos);
  };

  /** Edit → Paste: reads the clipboard natively (the same read Ctrl+V uses). */
  const pasteFromMenu = async () => {
    const text = await DocxKeys?.clipboardText().catch(() => null);
    if (text == null || text === '') {
      setStatus('The clipboard is empty.');
      return;
    }
    paste(text);
  };

  const deleteFromMenu = () => {
    if (!typing && selection) {
      remove();
    } else {
      nothingSelected();
    }
  };

  /** Keys the native listener caught (DocxKeysModule). Ctrl and Cmd both work. */
  const onKey = (e: KeyPress) => {
    switch (e.key) {
      case 'DEL_FWD':
        forwardDelete();
        return;
      case 'TAB':
        paste('\t');
        return;
      case 'LEFT':
      case 'RIGHT':
      case 'UP':
      case 'DOWN':
      case 'HOME':
      case 'END':
        move(e.key, e.shift);
        return;
      case 'Z':
        if (e.shift) {
          redo();
        } else {
          undo();
        }
        return;
      case 'Y':
        redo();
        return;
      case 'B':
      case 'I':
      case 'U':
        if (!typing && selection) {
          format(e.key === 'B' ? 'b' : e.key === 'I' ? 'i' : 'u', `${e.key} shortcut`);
        }
        return;
      case 'C':
      case 'X':
        copy(e.key === 'X');
        return;
      case 'V':
        if (e.text !== undefined) {
          paste(e.text);
        }
        return;
    }
  };
  const onKeyRef = useRef(onKey);
  onKeyRef.current = onKey;

  useEffect(() => {
    const sub = DeviceEventEmitter.addListener('DocxKey', (e: KeyPress) => onKeyRef.current(e));
    return () => sub.remove();
  }, []);

  // The listener goes on the invisible field whenever a new one mounts (a document opens).
  useEffect(() => {
    const tag = findNodeHandle(keys.current);
    if (!doc || readOnly || tag === null || !DocxKeys) {
      return;
    }
    DocxKeys.attach(tag)
      .then(r => log(`keys: ${r}`))
      .catch(error => log(`keys attach failed: ${errorText(error)}`));
  }, [doc, readOnly]);

  const undo = () => {
    flushTyping();
    setEdits(e => ({...e, cursor: Math.max(0, e.cursor - 1)}));
    setCaret(null);
    setSelection(null);
  };
  const redo = () => {
    flushTyping();
    setEdits(e => ({...e, cursor: Math.min(e.steps.length, e.cursor + 1)}));
    setCaret(null);
    setSelection(null);
  };

  /**
   * Save: over the document itself. The version on disk is backed up first (the last five
   * are kept, File > Previous versions), and if the file changed since DOCX read it — synced
   * from elsewhere, say — it is not overwritten: a copy is saved instead.
   */
  const save = async () => {
    if (!doc || !Docx || !doc.source || !doc.saveTo) {
      return;
    }
    // Text being typed is part of what is saved.
    const ops = [...applied, ...pending];
    const cursor = edits.cursor + (pending.length > 0 ? 1 : 0);
    flushTyping();
    setBusy(true);
    setStatus('Saving…');
    try {
      if (!(await ensureFileWritePermission())) {
        setStatus('Saving needs file write permission.');
        return;
      }
      const onDisk = await Docx.fileStamp(doc.saveTo);
      if (onDisk !== diskStamp) {
        const copy = await Docx.copyName(doc.saveTo);
        const res = await Docx.save(doc.source, ops, copy, expectedTexts(doc.blocks, ops));
        setStatus(`The document changed outside SNScribe since it was opened, so it was not overwritten. Your version was saved as ${res.name}.`);
        log(`save conflict: ${doc.saveTo} was ${diskStamp}, now ${onDisk}; saved ${res.dest}`);
        return;
      }
      if (!isNew || saved.dest) {
        await Docx.backup(doc.saveTo, docKey(doc.saveTo));
      }
      const res = await Docx.save(doc.source, ops, doc.saveTo, expectedTexts(doc.blocks, ops));
      setDiskStamp(await Docx.fileStamp(doc.saveTo));
      setSaved({cursor, dest: res.dest});
      setDiscardArmed(false);
      setStatus(`Saved ${res.name}.`);
    } catch (error) {
      setStatus(`Not saved: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  /** Save a copy: a new <name>-edited.docx beside the document, which stays as it is. */
  const saveCopy = async () => {
    if (!doc || !Docx || !doc.source || !doc.saveTo) {
      return;
    }
    const ops = [...applied, ...pending];
    flushTyping();
    setBusy(true);
    try {
      if (!(await ensureFileWritePermission())) {
        setStatus('Saving needs file write permission.');
        return;
      }
      const res = await Docx.save(doc.source, ops, await Docx.copyName(doc.saveTo), expectedTexts(doc.blocks, ops));
      setStatus(`Saved a copy as ${res.name}. The document itself is unchanged.`);
    } catch (error) {
      setStatus(`Copy not saved: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  // ---------------------------------------------------------------- recovery, versions, recent

  // Unsaved edits are recorded a moment after each change; nothing is kept once saved.
  useEffect(() => {
    if (!doc?.saveTo || !Docx) {
      return;
    }
    const key = `recovery-${docKey(doc.saveTo)}`;
    const t = setTimeout(() => {
      const rec = recoveryFor(doc.saveTo!, diskStamp, edits.steps, edits.cursor, saved.cursor, pending);
      if (rec) {
        Docx?.store(key, JSON.stringify(rec));
      } else {
        Docx?.forget(key);
      }
    }, 800);
    return () => clearTimeout(t);
  }, [doc, edits, saved, pending, diskStamp]);

  const restoreRecovery = () => {
    if (recovery) {
      setEdits({steps: recovery.steps, cursor: recovery.steps.length});
      setSaved({cursor: 0, dest: null});
      setStatus('Unsaved changes restored. Save to keep them.');
    }
    setRecovery(null);
    setMenu(null);
  };

  const discardRecovery = () => {
    if (doc?.saveTo) {
      Docx?.forget(`recovery-${docKey(doc.saveTo)}`);
    }
    setRecovery(null);
    setMenu(null);
  };

  const showVersions = async () => {
    if (!doc?.saveTo) {
      return;
    }
    setVersions(await Docx!.backups(docKey(doc.saveTo)));
    setMenu('versions');
  };

  /** Puts a previous version back as the document (the current one is backed up first), and reopens it. */
  const restoreVersion = async (from: string) => {
    if (!doc?.saveTo || !mayDiscard('the version')) {
      return;
    }
    setBusy(true);
    try {
      await Docx!.backup(doc.saveTo, docKey(doc.saveTo));
      await Docx!.copyOver(from, doc.saveTo);
      Docx!.forget(`recovery-${docKey(doc.saveTo)}`);
      await openPath(doc.saveTo);
      setStatus('Previous version restored. The version it replaced is in Previous versions too.');
    } catch (error) {
      setStatus(`Not restored: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  const openRecent = async (r: RecentDoc) => {
    if (!mayDiscard('the document')) {
      return;
    }
    if (!(await Docx!.fileStamp(r.path))) {
      setRecent(list => {
        const next = list.filter(x => x.path !== r.path);
        Docx?.store('recent', JSON.stringify(next));
        return next;
      });
      setStatus(`${r.name} is no longer there; removed from Recent.`);
      return;
    }
    setBusy(true);
    try {
      await openPath(r.path);
    } catch (error) {
      setStatus(`Could not open: ${errorText(error)}`);
    } finally {
      setBusy(false);
    }
  };

  // Settings are remembered between sessions.
  useEffect(() => {
    Docx?.store('settings', JSON.stringify({scaleAt, newFolder, lastName, author, marginOpen: showMargin, inkColor, citeStyle, spell: spellOn}));
  }, [scaleAt, newFolder, lastName, author, showMargin, inkColor, citeStyle, spellOn]);

  // The page follows the caret when it moves (typing, arrows) — not when the page is turned.
  const followCaret = useRef(false);
  useEffect(() => {
    followCaret.current = true;
  }, [caretAt?.para, caretAt?.offset]);

  useEffect(() => {
    const at = caretAt;
    if (!followCaret.current || !caretBox || caretBox.key !== pageKey || !at) {
      return;
    }
    followCaret.current = false;
    const bottom = brk.kind === 'at' ? brk.top : anchor.offset + pageH;
    if (caretBox.top >= bottom - 1 && brk.kind === 'at') {
      const to = anchorAfter(anchor, boxesNow(), brk);
      if (to && to.block < blocks.length) {
        setHistory(h => [...h, anchor]);
        setAnchor(to);
      }
    } else if (caretBox.top < anchor.offset - 1) {
      const i = blocks.findIndex(b => b.type === 'p' && b.index === at.para);
      if (i >= 0) {
        setAnchor(pageStartOf({block: i, offset: 0}));
      }
    }
    // Only a new caret box (for the moved caret) triggers this.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [caretBox]);

  // Where to draw the caret: asked of the paragraph's own layout once the page is measured.
  useEffect(() => {
    const at = caretAt;
    if (!at || brk.kind === 'pending' || !DocxText) {
      return;
    }
    const i = window.findIndex(b => b.type === 'p' && b.index === at.para);
    if (i < 0 && followCaret.current) {
      // The caret moved off this stretch of the document: go to its paragraph.
      followCaret.current = false;
      const bi = blocks.findIndex(b => b.type === 'p' && b.index === at.para);
      if (bi >= 0) {
        setHistory([]);
        setAnchor(pageStartOf({block: bi, offset: 0}));
      }
      return;
    }
    const f = frames.current[i];
    const tag = i >= 0 ? findNodeHandle(textRefs.current[i] ?? null) : null;
    if (i < 0 || !f || tag === null) {
      return;
    }
    const key = pageKey;
    // Typing changes the paragraph's layout: ask once it has been laid out again.
    const timer = setTimeout(() => {
      const fr = frames.current[i] ?? f;
      const tf = textFrames.current[i] ?? {x: 0, y: 0};
      DocxText?.caretRect(tag, toShown(window[i] as ParagraphBlock, at.offset)).then(r => {
        if (r.error !== undefined || measuredFor.current !== key) {
          return;
        }
        const scale = PixelRatio.get();
        setCaretBox({key, x: fr.x + tf.x + r.x! / scale, top: fr.top + tf.y + r.top! / scale, height: (r.bottom! - r.top!) / scale});
      });
    }, 40);
    return () => clearTimeout(timer);
    // frames/refs are read at call time; brk marks "measured".
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [caretAt?.para, caretAt?.offset, brk, pageKey]);

  // ---------------------------------------------------------------- render

  const visible = brk.kind === 'at' ? Math.max(0, brk.top - anchor.offset) : pageH;

  /** Where on the page (y) the text offset of window block i is, from its measured lines. */
  const yOf = (i: number, offset: number): number | null => {
    const b = window[i];
    const f = frames.current[i];
    const l = lines.current[i];
    if (!f || b?.type !== 'p') {
      return null;
    }
    const tf = textFrames.current[i] ?? {x: 0, y: 0};
    const shown = toShown(b, offset);
    let pos = 0;
    let y = 0;
    for (const line of l ?? []) {
      y = line.y;
      if (shown < pos + (line.len ?? 0)) {
        break;
      }
      pos += line.len ?? 0;
    }
    return f.top + tf.y + y - anchor.offset;
  };

  /** The folded margin: a strip with how many items this page has, and a button to open it. */
  const marginStrip = () => {
    let n = 0;
    window.forEach(b => {
      if (b.type === 'p') {
        n += (b.revs?.length ?? 0) + (b.marks ?? []).filter(m => m.kind === 'start').length + b.runs.filter(r => r.obj === 'ink').length;
      }
    });
    return (
      <Pressable
        onPress={once('open-margin', () => setShowMargin(true))}
        style={[styles.strip, {height: pageH, left: PAD + textW + MARGIN_GAP, top: PAD}]}>
        <Text allowFontScaling={false} style={styles.stripArrow}>
          {'‹'}
        </Text>
        {n > 0 ? (
          <Text allowFontScaling={false} style={styles.stripCount}>
            {String(n)}
          </Text>
        ) : null}
      </Pressable>
    );
  };

  /** The margin: a card beside the line of each note, comment and tracked change on this page. */
  const marginColumn = () => {
    type Card = {
      key: string;
      y: number;
      para: number;
      id: string;
      kind: 'ins' | 'del' | 'comment' | 'ink';
      author: string;
      text: string;
      move?: boolean;
      replies?: number;
      h?: number;
      src?: string;
    };
    const cards: Card[] = [];
    const inkW = MARGIN_W - 20;
    window.forEach((b, i) => {
      if (b.type !== 'p') {
        return;
      }
      let offset = 0;
      for (const r of b.runs) {
        if (r.obj === 'ink' && r.ink) {
          const y = yOf(i, offset);
          if (y !== null && y >= -4 && y < visible) {
            cards.push({key: `n:${r.ink}`, y: Math.max(0, y), para: b.index, id: r.ink, kind: 'ink', author: '', text: '', h: 120, src: inks[r.ink]});
          }
        }
        offset += r.t.length;
      }
    });
    const threads = new Set(comments.filter(c => !c.parent).map(c => c.id));
    window.forEach((b, i) => {
      if (b.type !== 'p') {
        return;
      }
      for (const m of b.marks ?? []) {
        // A thread's card sits at its start (or, with no range, at its reference mark).
        const hasStart = (b.marks ?? []).some(x => x.id === m.id && x.kind === 'start');
        if (!threads.has(m.id) || (m.kind !== 'start' && !(m.kind === 'ref' && !hasStart))) {
          continue;
        }
        const c = comments.find(x => x.id === m.id)!;
        const y = yOf(i, m.at);
        if (y !== null && y >= -4 && y < visible) {
          cards.push({
            key: `c:${m.id}`,
            y: Math.max(0, y),
            para: b.index,
            id: m.id,
            kind: 'comment',
            author: c.author,
            text: c.text || (c.pictures ? '✎ handwritten note' : ''),
            replies: threadIds(comments, m.id).length - 1,
          });
        }
      }
      for (const v of b.revs ?? []) {
        let at = v.at ?? 0;
        let text = (v.runs ?? []).map(r => r.t).join('');
        if (v.kind === 'ins') {
          let pos = 0;
          let first = -1;
          text = '';
          for (const r of b.runs) {
            if (r.rv === v.id) {
              first = first < 0 ? pos : first;
              text += r.t;
            }
            pos += r.t.length;
          }
          at = Math.max(0, first);
        }
        const y = yOf(i, at);
        if (y !== null && y >= -4 && y < visible) {
          cards.push({key: `${b.index}:${v.id}:${v.kind}`, y: Math.max(0, y), para: b.index, id: v.id, kind: v.kind, author: v.author, text, move: v.move});
        }
      }
    });
    cards.sort((a, b) => a.y - b.y);
    // Stacked downward where they would overlap; those that no longer fit are counted.
    let bottom = 0;
    const placed: Array<Card & {top: number}> = [];
    let hidden = 0;
    for (const c of cards) {
      const top = Math.max(c.y, bottom);
      const h = c.h ?? CARD_H;
      if (top + h > pageH - 84) {
        hidden++;
        continue;
      }
      placed.push({...c, top});
      bottom = top + h + 6;
    }
    return (
      <View style={[styles.margin, {height: pageH, left: PAD + textW + MARGIN_GAP, top: PAD}]}>
        {placed.map(c =>
          c.kind === 'ink' ? (
            <Pressable
              key={c.key}
              onPress={once(`note:${c.id}`, () => {
                setNoteOpen(c.id);
                setMenu('note');
              })}
              style={[styles.inkCard, {top: c.top, height: c.h}]}>
              {c.src ? (
                <Image source={{uri: `file://${c.src}`}} resizeMode="contain" style={{width: inkW, height: (c.h ?? CARD_H) - 8}} />
              ) : (
                <Text allowFontScaling={false} style={styles.cardBody}>
                  {'✎ Handwritten note'}
                </Text>
              )}
            </Pressable>
          ) : c.kind === 'comment' ? (
            <Pressable key={c.key} onPress={once(`thread:${c.id}`, () => openThread(c.id))} style={[styles.card, styles.commentCard, {top: c.top}]}>
              <View style={styles.cardText}>
                <Text allowFontScaling={false} style={styles.cardHead} numberOfLines={1}>
                  {`Comment · ${c.author || 'Unknown'}${c.replies ? `  ·  ${c.replies} repl${c.replies === 1 ? 'y' : 'ies'}` : ''}`}
                </Text>
                <Text allowFontScaling={false} style={styles.cardBody} numberOfLines={2}>
                  {c.text}
                </Text>
              </View>
            </Pressable>
          ) : (
          <View key={c.key} style={[styles.card, {top: c.top}]}>
            <View style={styles.cardText}>
              <Text allowFontScaling={false} style={styles.cardHead} numberOfLines={1}>
                {`${c.move ? (c.kind === 'ins' ? 'Moved here' : 'Moved away') : c.kind === 'ins' ? 'Inserted' : 'Deleted'} · ${c.author || 'Unknown'}`}
              </Text>
              <Text allowFontScaling={false} style={[styles.cardBody, c.kind === 'del' ? styles.cardDeleted : styles.cardInserted]} numberOfLines={2}>
                {c.text.replace(/\ufffc/g, '◇')}
              </Text>
            </View>
            <Pressable onPress={once(`acc:${c.key}`, () => review(c.para, c.id, true))} style={styles.cardButton}>
              <Text allowFontScaling={false} style={styles.cardButtonText}>
                {'✓'}
              </Text>
            </Pressable>
            <Pressable onPress={once(`rej:${c.key}`, () => review(c.para, c.id, false))} style={styles.cardButton}>
              <Text allowFontScaling={false} style={styles.cardButtonText}>
                {'✗'}
              </Text>
            </Pressable>
          </View>
          ),
        )}
        {hidden > 0 ? (
          <Text allowFontScaling={false} style={[styles.cardMore, {top: pageH - 76}]}>
            {`+${hidden} more on this page (turn the page or hide some by reviewing)`}
          </Text>
        ) : null}
        {doc?.otherRevisions ? (
          <Text allowFontScaling={false} style={[styles.cardMore, {top: placed.length ? pageH - 100 : 0}]} numberOfLines={2}>
            {`${doc.otherRevisions} formatting or paragraph change${doc.otherRevisions === 1 ? '' : 's'} can't be reviewed here yet; they stay as they are.`}
          </Text>
        ) : null}
        <Pressable onPress={once('fold-margin', () => setShowMargin(false))} style={[styles.foldButton, {top: pageH - 44}]}>
          <Text allowFontScaling={false} style={styles.foldText}>
            {'›  Fold panel'}
          </Text>
        </Pressable>
      </View>
    );
  };
  const progress = doc && blocks.length > 0 ? Math.round((anchor.block / blocks.length) * 100) : 0;
  // From the committed text: typed words reach Contents when typing ends, not on every key.
  const headings = useMemo(() => outline(committed), [committed]);

  // A pen tap sometimes arrives as two presses; a second press of the same control this soon is ignored.
  const lastPress = useRef<{key: string; at: number}>({key: '', at: 0});
  const once = (key: string, action: () => void) => () => {
    const now = Date.now();
    if (lastPress.current.key === key && now - lastPress.current.at < 400) {
      return;
    }
    lastPress.current = {key, at: now};
    action();
  };

  /** A plain button: closes any open menu, then acts. */
  const button = (label: string, action: () => void, disabled = false, extra?: object) => (
    <Pressable
      key={label}
      disabled={disabled || busy}
      onPress={once(label, () => {
        setMenu(null);
        action();
      })}
      style={[styles.button, extra, disabled || busy ? styles.disabled : null]}>
      <Text allowFontScaling={false} style={styles.buttonText}>
        {label}
      </Text>
    </Pressable>
  );

  /**
   * A button inside a panel that must stay open (Search, Next, Reply …). `button` closes
   * the open menu first — right for the toolbar, wrong in a panel: its results then arrived
   * after it had closed (CT: Search "just goes back to the document").
   */
  const panelButton = (label: string, action: () => void, disabled = false) => (
    <Pressable
      key={label}
      disabled={disabled || busy}
      onPress={once(`panel:${label}`, action)}
      style={[styles.button, disabled || busy ? styles.disabled : null]}>
      <Text allowFontScaling={false} style={styles.buttonText}>
        {label}
      </Text>
    </Pressable>
  );

  /** A menu button: opens its menu (never toggles it shut — closing is by choosing, or tapping elsewhere). */
  const menuButton = (label: string, which: Menu) => (
    <Pressable
      key={which}
      disabled={busy}
      onLayout={e => {
        menuXs.current[which] = e.nativeEvent.layout.x;
      }}
      onPress={once(`menu:${which}`, () => {
        setMenuX((menuXs.current[which] ?? 0) - (which === 'file' ? 0 : toolScroll.current));
        setMenu(which);
      })}
      style={[styles.button, menu === which ? styles.buttonOpen : null, busy ? styles.disabled : null]}>
      <Text allowFontScaling={false} style={[styles.buttonText, menu === which ? styles.buttonOpenText : null]}>
        {`${label} ▾`}
      </Text>
    </Pressable>
  );
  const menuXs = useRef<Partial<Record<Menu, number>>>({});
  const toolScroll = useRef(0);

  /** One row of a drop-down menu. */
  const item = (label: string, action: () => void, style?: object) => (
    <Pressable
      key={label}
      onPress={once(`item:${label}`, () => {
        setMenu(null);
        action();
      })}
      style={styles.menuItem}>
      <Text allowFontScaling={false} style={[styles.menuText, style]} numberOfLines={1}>
        {label}
      </Text>
    </Pressable>
  );

  const menuBody = (): React.JSX.Element | null => {
    switch (menu) {
      case 'file':
        return (
          <>
            {item('New…', startNew)}
            {item('New from template…', startFromTemplate)}
            {item('Open…', open)}
            {recent.length ? item('Recent…', () => setMenu('recent')) : null}
            {doc && !readOnly ? item('Page setup…', () => setMenu('page')) : null}
            {doc && !readOnly ? item('Header, footer & page numbers…', openHeaderForm) : null}
            {doc && !readOnly ? item('Paper format (APA, MLA, Chicago)…', () => setMenu('preset')) : null}
            {doc && !readOnly ? item('Save', save) : null}
            {doc && !readOnly ? item('Save a copy', saveCopy) : null}
            {doc && !readOnly ? item('Previous versions…', showVersions) : null}
            {item('Close', close)}
          </>
        );
      case 'recent':
        return (
          <>
            {recent.map(r => item(`${r.name.replace(/\.docx$/i, '')}  ·  ${r.path.split('/').slice(-2, -1)[0] ?? ''}`, () => openRecent(r)))}
          </>
        );
      case 'versions':
        return (
          <>
            {versions.length === 0 ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.folderPath]}>
                {'No earlier versions yet. Each Save keeps the version it replaces.'}
              </Text>
            ) : null}
            {versions.map(v =>
              item(`${new Date(v.time).toLocaleString()}  ·  ${Math.max(1, Math.round(v.bytes / 1024))} KB`, () => restoreVersion(v.path)),
            )}
          </>
        );
      case 'recover':
        return recovery ? (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {`This document has unsaved changes from ${new Date(recovery.time).toLocaleString()} (${recovery.steps.length} edit${recovery.steps.length === 1 ? '' : 's'}).`}
            </Text>
            <View style={styles.row}>
              {button('Restore', restoreRecovery)}
              {button('Discard', discardRecovery)}
            </View>
          </View>
        ) : null;
      case 'para': {
        const cur = paraShown();
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        const line = cur?.line !== undefined && (cur.lineRule ?? 'auto') === 'auto' ? cur.line : undefined;
        const pt = (tw?: number) => (tw === undefined ? undefined : Math.round(tw / 20));
        return (
          <View style={styles.paraMenu}>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Alignment'}
            </Text>
            <View style={styles.chips}>
              {(['left', 'center', 'right', 'justify'] as const).map(a =>
                chip(a[0].toUpperCase() + a.slice(1), cur?.align === a, () => paraTool({align: a}, `align ${a}`)),
              )}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Line spacing'}
            </Text>
            <View style={styles.chips}>
              {[
                ['1.0', 240],
                ['1.15', 276],
                ['1.5', 360],
                ['2.0', 480],
              ].map(([label, v]) => chip(String(label), line === v, () => paraTool({line: Number(v), lineRule: 'auto'}, `line ${label}`)))}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Space before (pt)'}
            </Text>
            <View style={styles.chips}>
              {[0, 6, 12, 18].map(v => chip(String(v), pt(cur?.before) === v, () => paraTool({before: v * 20}, `before ${v}`)))}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Space after (pt)'}
            </Text>
            <View style={styles.chips}>
              {[0, 6, 8, 12].map(v => chip(String(v), pt(cur?.after) === v, () => paraTool({after: v * 20}, `after ${v}`)))}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Indent'}
            </Text>
            <View style={styles.chips}>
              {chip('None', !cur?.first, () => paraTool({first: 0}, 'no indent'))}
              {chip('First line 0.5″', (cur?.first ?? 0) > 0, () => paraTool({first: 720}, 'first-line indent'))}
              {chip('Hanging 0.5″', (cur?.first ?? 0) < 0, () => paraTool({first: -720}, 'hanging indent'))}
            </View>
            <View style={styles.chips}>
              {chip(cur?.pb ? '✓ Starts on a new page' : 'Start on a new page', !!cur?.pb, () =>
                paraTool({pb: !cur?.pb}, cur?.pb ? 'no page break' : 'page break before'),
              )}
            </View>
            <View style={styles.chips}>
              {chip('Insert page break here', false, () => {
                setMenu(null);
                insertPageBreak();
              })}
              {chip('Done', false, () => setMenu(null))}
            </View>
          </View>
        );
      }
      case 'page': {
        const pg = pageAfter(doc?.page, applied);
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        const setPage = (props: Omit<Extract<Op, {op: 'page'}>, 'op' | 'para'>, label: string) =>
          commit([{op: 'page', para: -1, ...props}], label);
        const margin = pg.top === pg.bottom && pg.left === pg.right && pg.top === pg.left ? pg.top : null;
        const short = Math.min(pg.width, pg.height);
        const paper = Math.abs(short - 12240) < 60 ? 'letter' : Math.abs(short - 11906) < 60 ? 'a4' : 'other';
        const size = (w: number, h: number) => (pg.landscape ? {width: h, height: w} : {width: w, height: h});
        return (
          <View style={styles.paraMenu}>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Margins (all sides)'}
            </Text>
            <View style={styles.chips}>
              {[
                ['0.5″', 720],
                ['0.75″', 1080],
                ['1″', 1440],
                ['1.25″', 1800],
                ['1.5″', 2160],
              ].map(([label, v]) =>
                chip(String(label), margin === v, () => setPage({top: Number(v), right: Number(v), bottom: Number(v), left: Number(v)}, `margins ${label}`)),
              )}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Paper'}
            </Text>
            <View style={styles.chips}>
              {chip('Letter', paper === 'letter', () => setPage(size(12240, 15840), 'Letter paper'))}
              {chip('A4', paper === 'a4', () => setPage(size(11906, 16838), 'A4 paper'))}
            </View>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {'Orientation'}
            </Text>
            <View style={styles.chips}>
              {chip('Portrait', !pg.landscape, () =>
                pg.landscape && setPage({landscape: false, width: Math.min(pg.width, pg.height), height: Math.max(pg.width, pg.height)}, 'portrait'),
              )}
              {chip('Landscape', pg.landscape, () =>
                !pg.landscape && setPage({landscape: true, width: Math.max(pg.width, pg.height), height: Math.min(pg.width, pg.height)}, 'landscape'),
              )}
            </View>
            <Text allowFontScaling={false} style={styles.panelNote}>
              {'Margins and paper shape the printed Word document; on the Supernote the text always fits the screen.'}
            </Text>
            <View style={styles.chips}>{chip('Done', false, () => setMenu(null))}</View>
          </View>
        );
      }
      case 'preset':
        return (
          <View>
            <View style={[styles.row, styles.presetName]}>
              <Text allowFontScaling={false} style={styles.menuText}>
                {'Last name (MLA)'}
              </Text>
              <TextInput
                style={styles.nameInput}
                value={lastName}
                onChangeText={setLastName}
                allowFontScaling={false}
                autoCorrect={false}
                returnKeyType="done"
              />
            </View>
            {PAPER_FORMATS.map(f => (
              <Pressable
                key={f.id}
                onPress={once(`preset:${f.id}`, () => {
                  setMenu(null);
                  flushTyping();
                  commit(presetOps(blocks, f.id, lastName), f.name);
                  setStatus(
                    f.id === 'mla' && !lastName.trim()
                      ? `Formatted as ${f.name}. Add your last name above to put it before the page number.`
                      : `Formatted as ${f.name}. Undo puts it back as it was.`,
                  );
                })}
                style={styles.menuItem}>
                <Text allowFontScaling={false} style={[styles.menuText, styles.menuAction]}>
                  {f.name}
                </Text>
                <Text allowFontScaling={false} style={styles.presetSummary}>
                  {f.summary}
                </Text>
              </Pressable>
            ))}
          </View>
        );
      case 'header': {
        if (!hfForm) {
          return null;
        }
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        const set = (kind: 'header' | 'footer', change: Partial<HfLine>) =>
          setHfForm(f => (f ? {...f, [kind]: {...f[kind], ...change}} : f));
        const section = (kind: 'header' | 'footer', title: string, other?: boolean) => (
          <View key={kind}>
            <Text allowFontScaling={false} style={styles.paraLabel}>
              {title}
            </Text>
            <TextInput
              style={styles.nameInput}
              value={hfForm[kind].text}
              onChangeText={text => set(kind, {text})}
              allowFontScaling={false}
              placeholder={kind === 'header' ? 'e.g. your last name (optional)' : 'optional'}
              returnKeyType="done"
            />
            <View style={[styles.chips, styles.hfChips]}>
              {(['left', 'center', 'right'] as const).map(a =>
                chip(a[0].toUpperCase() + a.slice(1), hfForm[kind].align === a, () => set(kind, {align: a})),
              )}
              {chip(hfForm[kind].page ? '✓ Page number' : 'Page number', hfForm[kind].page, () => set(kind, {page: !hfForm[kind].page}))}
            </View>
            {other ? (
              <Text allowFontScaling={false} style={styles.panelNote}>
                {`This ${kind} has a logo or table. It stays; only the text and page number line changes.`}
              </Text>
            ) : null}
          </View>
        );
        return (
          <View style={styles.paraMenu}>
            {section('header', 'Header (top of every page)', headerNow()?.other)}
            {section('footer', 'Footer (bottom of every page)', footerNow()?.other)}
            <Text allowFontScaling={false} style={styles.panelNote}>
              {'Only what you change is written. A different first page keeps its own. Word shows these; the Supernote page does not.'}
            </Text>
            <View style={styles.row}>
              {panelButton('Apply', applyHeader)}
              {button('Cancel', () => {
                setHfForm(null);
                setMenu(null);
              })}
            </View>
          </View>
        );
      }
      case 'cite': {
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        if (cite.setup) {
          return (
            <View style={styles.nameForm}>
              <Text allowFontScaling={false} style={styles.paraLabel}>
                {'Connect your Zotero library'}
              </Text>
              <Text allowFontScaling={false} style={styles.presetSummary}>
                {'On zotero.org: Settings → Security → Create new private key. Tick only “Allow library access” (read-only). The page shows your user ID (a number) above the key.'}
              </Text>
              <Text allowFontScaling={false} style={[styles.menuText, styles.findLabel]}>
                {'User ID'}
              </Text>
              <TextInput style={styles.nameInput} value={cite.userId} onChangeText={userId => setCite(c => ({...c, userId}))} keyboardType="number-pad" allowFontScaling={false} />
              <Text allowFontScaling={false} style={[styles.menuText, styles.findLabel]}>
                {zotero ? 'API key (leave empty to keep the saved one)' : 'API key'}
              </Text>
              <TextInput
                style={styles.nameInput}
                value={cite.apiKey}
                onChangeText={apiKey => setCite(c => ({...c, apiKey}))}
                autoCapitalize="none"
                autoCorrect={false}
                secureTextEntry
                allowFontScaling={false}
              />
              <View style={styles.row}>{panelButton('Load from file…', zoteroFromFile, cite.busy)}</View>
              <Text allowFontScaling={false} style={styles.presetSummary}>
                {'Load from file: a .txt with the key (and your user ID, if you like). Kept only on this Supernote, in the plugin’s private storage.'}
              </Text>
              {cite.message ? (
                <Text allowFontScaling={false} style={[styles.menuText, styles.citeMessage]}>
                  {cite.message}
                </Text>
              ) : null}
              <View style={styles.row}>
                {panelButton(cite.busy ? 'Checking…' : 'Save', saveZotero, cite.busy)}
                {panelButton('Cancel', () => (zotero ? setCite(c => ({...c, setup: false})) : setMenu(null)))}
              </View>
            </View>
          );
        }
        const chosen = cite.chosen;
        return (
          <View style={styles.nameForm}>
            <View style={styles.chips}>
              {CITE_STYLES.map(st =>
                chip(st.name, citeStyle === st.id, () => {
                  setCiteStyle(st.id);
                  setCite(c => ({...c, chosen: null}));
                  if (cite.results) {
                    searchCite(st.id);
                  }
                }),
              )}
            </View>
            <View style={styles.row}>
              <TextInput
                style={styles.nameInput}
                value={cite.query}
                onChangeText={query => setCite(c => ({...c, query}))}
                placeholder="Author, title or year"
                autoCorrect={false}
                allowFontScaling={false}
                returnKeyType="search"
                onSubmitEditing={() => searchCite()}
              />
              {panelButton(cite.busy ? '…' : 'Search', () => searchCite(), cite.busy || !cite.query.trim())}
            </View>
            {cite.message ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.citeMessage]}>
                {cite.message}
              </Text>
            ) : null}
            {!chosen
              ? (cite.results ?? []).map(r => (
                  <Pressable key={r.key} onPress={once(`src:${r.key}`, () => setCite(c => ({...c, chosen: r})))} style={styles.menuItem}>
                    <Text allowFontScaling={false} style={styles.cardHead} numberOfLines={1}>
                      {`${r.authors || 'No author'}${r.year ? ` (${r.year})` : ''}`}
                    </Text>
                    <Text allowFontScaling={false} style={styles.presetSummary} numberOfLines={2}>
                      {r.title}
                    </Text>
                  </Pressable>
                ))
              : null}
            {chosen ? (
              <View>
                <Text allowFontScaling={false} style={[styles.cardHead, styles.findLabel]} numberOfLines={2}>
                  {chosen.title}
                </Text>
                <View style={[styles.chips, styles.findLabel]}>
                  {chip('Parenthetical', !cite.narrative, () => setCite(c => ({...c, narrative: false})))}
                  {chip('Narrative', cite.narrative, () => setCite(c => ({...c, narrative: true})))}
                </View>
                <View style={[styles.row, styles.presetName]}>
                  <Text allowFontScaling={false} style={styles.menuText}>
                    {'Page (optional)'}
                  </Text>
                  <TextInput style={styles.nameInput} value={cite.page} onChangeText={page => setCite(c => ({...c, page}))} allowFontScaling={false} />
                </View>
                <Text allowFontScaling={false} style={styles.paraLabel}>
                  {'In the text'}
                </Text>
                <Text allowFontScaling={false} style={styles.menuText}>
                  {inText(chosen, citeStyle, cite.narrative, cite.page)}
                </Text>
                <Text allowFontScaling={false} style={styles.paraLabel}>
                  {`In ${CITE_STYLES.find(x => x.id === citeStyle)!.heading}`}
                </Text>
                <Text allowFontScaling={false} style={styles.presetSummary}>
                  {htmlToPieces(chosen.bibHtml).map((pc, i) => (
                    <Text key={i} style={pc.i ? styles.italicText : null}>
                      {pc.t}
                    </Text>
                  ))}
                </Text>
                <View style={styles.row}>
                  {panelButton('Insert', insertCitation)}
                  {panelButton('Back', () => setCite(c => ({...c, chosen: null})))}
                </View>
              </View>
            ) : null}
            <View style={styles.row}>
              {button('Close', () => setMenu(null))}
              {panelButton('Zotero settings', () => setCite(c => ({...c, setup: true, userId: zotero?.userId ?? ''})))}
            </View>
          </View>
        );
      }
      case 'quotes': {
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        const msg = qp.message ? (
          <Text allowFontScaling={false} style={[styles.menuText, styles.citeMessage]}>
            {qp.message}
          </Text>
        ) : null;
        const q = qp.chosen;
        const head = q ? (
          <View>
            <Text allowFontScaling={false} style={styles.menuText} numberOfLines={3}>
              {`“${q.text}”`}
            </Text>
            <Text allowFontScaling={false} style={styles.presetSummary} numberOfLines={1}>
              {`${fileName(q.path)} · p. ${q.page}`}
            </Text>
          </View>
        ) : null;
        if (qp.mode === 'list' || !q) {
          return (
            <View style={styles.nameForm}>
              {msg}
              {qp.quotes.map(x => (
                <Pressable key={x.id} onPress={once(`quote:${x.id}`, () => chooseQuote(x))} style={styles.menuItem}>
                  <Text allowFontScaling={false} style={styles.menuText} numberOfLines={2}>
                    {`“${x.text}”`}
                  </Text>
                  <Text allowFontScaling={false} style={styles.presetSummary} numberOfLines={1}>
                    {`${fileName(x.path)} · p. ${x.page}${qp.sources[x.path] ? '' : ' · source not set yet'}`}
                  </Text>
                </Pressable>
              ))}
              <View style={styles.row}>{button('Close', () => setMenu(null))}</View>
            </View>
          );
        }
        if (qp.mode === 'source') {
          return (
            <View style={styles.nameForm}>
              {head}
              <Text allowFontScaling={false} style={styles.paraLabel}>
                {`Where do the details for ${fileName(q.path)} come from? (Asked once per file.)`}
              </Text>
              <View style={[styles.row, styles.presetName]}>
                <Text allowFontScaling={false} style={styles.menuText}>
                  {'DOI'}
                </Text>
                <TextInput style={styles.nameInput} value={qp.doi} onChangeText={doi => setQp(x => ({...x, doi}))} autoCapitalize="none" autoCorrect={false} allowFontScaling={false} placeholder="10.xxxx/…" />
                {panelButton('Look up', sourceFromDoi, qp.busy || !qp.doi.trim())}
              </View>
              <View style={styles.chips}>
                {isEpub(q.path) ? chip('Use the EPUB’s details', false, sourceFromEpub) : null}
                {chip('Find in Zotero', false, () => setQp(x => ({...x, mode: 'zotero', results: [], message: ''})))}
                {chip('Type the details', false, () => setQp(x => ({...x, mode: 'form', message: ''})))}
              </View>
              {msg}
              <View style={styles.row}>{panelButton('Back', () => setQp(x => ({...x, mode: 'list', chosen: null, message: ''})))}</View>
            </View>
          );
        }
        if (qp.mode === 'zotero') {
          return (
            <View style={styles.nameForm}>
              {head}
              <View style={styles.row}>
                <TextInput style={styles.nameInput} value={qp.query} onChangeText={query => setQp(x => ({...x, query}))} autoCorrect={false} allowFontScaling={false} returnKeyType="search" onSubmitEditing={searchQuoteZotero} />
                {panelButton(qp.busy ? '…' : 'Search', searchQuoteZotero, qp.busy || !qp.query.trim())}
              </View>
              {msg}
              {qp.results.map(r => (
                <Pressable key={r.key} onPress={once(`qz:${r.key}`, () => takeSource({kind: 'zotero', key: r.key}))} style={styles.menuItem}>
                  <Text allowFontScaling={false} style={styles.cardHead} numberOfLines={1}>
                    {`${r.authors || 'No author'}${r.year ? ` (${r.year})` : ''}`}
                  </Text>
                  <Text allowFontScaling={false} style={styles.presetSummary} numberOfLines={2}>
                    {r.title}
                  </Text>
                </Pressable>
              ))}
              <View style={styles.row}>{panelButton('Back', () => setQp(x => ({...x, mode: 'source', message: ''})))}</View>
            </View>
          );
        }
        if (qp.mode === 'form') {
          const f = qp.form;
          const set = (k: keyof typeof f) => (v: string) => setQp(x => ({...x, form: {...x.form, [k]: v}}));
          const field = (label: string, k: keyof typeof f, multiline = false) => (
            <View key={k}>
              <Text allowFontScaling={false} style={[styles.menuText, styles.findLabel]}>
                {label}
              </Text>
              <TextInput
                style={[styles.nameInput, multiline ? styles.commentInput : null]}
                value={String(f[k])}
                onChangeText={set(k)}
                multiline={multiline}
                autoCorrect={false}
                allowFontScaling={false}
              />
            </View>
          );
          return (
            <View style={styles.nameForm}>
              <View style={styles.chips}>
                {chip('Article', f.type === 'article', () => set('type')('article'))}
                {chip('Book', f.type === 'book', () => set('type')('book'))}
                {chip('Other (web, report)', f.type === 'other', () => set('type')('other'))}
              </View>
              {field('Authors — one per line: Last, First', 'authors', true)}
              {field('Year', 'year')}
              {field('Title', 'title')}
              {f.type !== 'book' ? field(f.type === 'article' ? 'Journal' : 'Website or larger work', 'container') : null}
              {f.type === 'article' ? field('Volume', 'volume') : null}
              {f.type === 'article' ? field('Issue', 'issue') : null}
              {f.type === 'article' ? field('Pages (e.g. 29-40)', 'pages') : null}
              {f.type === 'book' ? field('Publisher', 'publisher') : null}
              {field('DOI or web address (optional)', 'link')}
              {msg}
              <View style={styles.row}>
                {panelButton('Use these details', () => {
                  const d = formDetails();
                  if (d) {
                    takeSource({kind: 'details', details: d});
                  }
                })}
                {panelButton('Back', () => setQp(x => ({...x, mode: 'source', message: ''})))}
              </View>
            </View>
          );
        }
        const src = qp.source;
        const long = q.text.trim().split(/\s+/).length >= 40;
        return (
          <View style={styles.nameForm}>
            {head}
            <View style={[styles.row, styles.presetName]}>
              <Text allowFontScaling={false} style={styles.menuText}>
                {'Page'}
              </Text>
              <TextInput style={styles.nameInput} value={qp.page} onChangeText={page => setQp(x => ({...x, page}))} allowFontScaling={false} />
            </View>
            {src ? (
              <>
                <Text allowFontScaling={false} style={styles.paraLabel}>
                  {long ? 'As a block quote, cited' : 'In the text'}
                </Text>
                <Text allowFontScaling={false} style={styles.menuText}>
                  {inText(src, citeStyle, false, qp.page)}
                </Text>
                <Text allowFontScaling={false} style={styles.paraLabel}>
                  {`In ${CITE_STYLES.find(x => x.id === citeStyle)!.heading}`}
                </Text>
                <Text allowFontScaling={false} style={styles.presetSummary}>
                  {htmlToPieces(src.bibHtml).map((pc, i) => (
                    <Text key={i} style={pc.i ? styles.italicText : null}>
                      {pc.t}
                    </Text>
                  ))}
                </Text>
              </>
            ) : null}
            {msg}
            <View style={styles.row}>
              {panelButton('Insert', insertQuote, !src)}
              {panelButton('Change source', () => setQp(x => ({...x, mode: 'source', message: ''})))}
            </View>
            <View style={styles.row}>
              {panelButton('Delete quote', () => {
                saveQuotes(qp.quotes.filter(x => x.id !== q.id));
                setQp(x => ({...x, mode: 'list', chosen: null, message: 'Quote deleted.'}));
              })}
              {panelButton('Back', () => setQp(x => ({...x, mode: 'list', chosen: null, message: ''})))}
            </View>
          </View>
        );
      }
      case 'spell': {
        const m = sp.items[sp.at];
        const ctx = m ? textOf(m.para) : '';
        const chip = (label: string, on: boolean, action: () => void) => (
          <Pressable key={label} onPress={once(`chip:${label}`, action)} style={[styles.chip, on ? styles.chipOn : null]}>
            <Text allowFontScaling={false} style={[styles.chipText, on ? styles.chipTextOn : null]}>
              {label}
            </Text>
          </Pressable>
        );
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.presetSummary}>
              {'English (US)'}
            </Text>
            {sp.message ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.citeMessage]}>
                {sp.message}
              </Text>
            ) : null}
            {m ? (
              <>
                <Text allowFontScaling={false} style={[styles.paraLabel, styles.findLabel]}>
                  {`“${m.word}”`}
                </Text>
                <Text allowFontScaling={false} style={styles.presetSummary} numberOfLines={2}>
                  {`…${ctx.slice(Math.max(0, m.start - 40), m.start)}`}
                  <Text style={styles.spellWord}>{ctx.slice(m.start, m.end)}</Text>
                  {`${ctx.slice(m.end, m.end + 40)}…`}
                </Text>
                <View style={[styles.chips, styles.findLabel]}>
                  {sp.suggestions.length
                    ? sp.suggestions.map(sg => chip(sg, sp.replace === sg, () => setSp(x => ({...x, replace: sg}))))
                    : (
                      <Text allowFontScaling={false} style={styles.presetSummary}>
                        {'No suggestions — type the right spelling below.'}
                      </Text>
                    )}
                </View>
                <TextInput style={styles.nameInput} value={sp.replace} onChangeText={replace => setSp(x => ({...x, replace}))} autoCorrect={false} autoCapitalize="none" allowFontScaling={false} />
                <View style={styles.row}>
                  {panelButton('Change', () => spellChange(false), !sp.replace.trim())}
                  {panelButton('Change all', () => spellChange(true), !sp.replace.trim())}
                </View>
                <View style={styles.row}>
                  {panelButton('Ignore', () => spellSkip())}
                  {panelButton('Ignore all', spellIgnoreAll)}
                  {panelButton('Add to dictionary', spellAdd)}
                </View>
                <View style={styles.row}>
                  {panelButton('‹ Prev', () => showSpell(sp.items, sp.at - 1), sp.at === 0)}
                  {panelButton('Next ›', () => showSpell(sp.items, sp.at + 1), sp.at >= sp.items.length - 1)}
                  {button('Close', () => setMenu(null))}
                </View>
              </>
            ) : (
              <View style={styles.row}>{button('Close', () => setMenu(null))}</View>
            )}
          </View>
        );
      }
      case 'comment':
        return commentForm ? (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText} numberOfLines={2}>
              {commentForm.quote ? `Comment on “${commentForm.quote.slice(0, 80)}”` : 'Comment at the caret'}
            </Text>
            <TextInput
              style={[styles.nameInput, styles.commentInput]}
              value={commentForm.text}
              onChangeText={text => setCommentForm(f => (f ? {...f, text} : f))}
              autoFocus
              multiline
              allowFontScaling={false}
              placeholder="Write or type the comment"
            />
            <View style={[styles.row, styles.presetName]}>
              <Text allowFontScaling={false} style={styles.menuText}>
                {'Your name'}
              </Text>
              <TextInput style={styles.nameInput} value={author} onChangeText={setAuthor} allowFontScaling={false} autoCorrect={false} />
            </View>
            <View style={styles.row}>
              {panelButton('Add comment', addComment)}
              {button('Cancel', () => {
                setCommentForm(null);
                setMenu(null);
              })}
            </View>
          </View>
        ) : null;
      case 'note': {
        const src = noteOpen ? inks[noteOpen] : undefined;
        return noteOpen ? (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.cardHead}>
              {'Handwritten note · in the right margin of the Word file'}
            </Text>
            {src ? <Image source={{uri: `file://${src}`}} resizeMode="contain" style={{width: MENU_W - 32, height: 300, marginVertical: 10}} /> : null}
            <View style={styles.row}>
              {button('Delete note', () => deleteNote(noteOpen))}
              {button('Close', () => {
                setNoteOpen(null);
                setMenu(null);
              })}
            </View>
          </View>
        ) : null;
      }
      case 'thread': {
        const ids = thread ? threadIds(comments, thread).map(String) : [];
        const items = comments.filter(c => ids.includes(c.id));
        return (
          <View style={styles.nameForm}>
            {items.map((c, i) => (
              <View key={c.id} style={i > 0 ? styles.threadReply : null}>
                <Text allowFontScaling={false} style={styles.cardHead}>
                  {`${c.author || 'Unknown'}${c.date ? `  ·  ${c.date.slice(0, 10)}` : ''}`}
                </Text>
                <Text allowFontScaling={false} style={styles.menuText}>
                  {c.text || (c.pictures ? '' : '(empty)')}
                </Text>
                {c.pictures ? (
                  <Text allowFontScaling={false} style={styles.panelNote}>
                    {'✎ A picture or handwritten note — it shows in Word.'}
                  </Text>
                ) : null}
              </View>
            ))}
            <TextInput
              style={[styles.nameInput, styles.commentInput]}
              value={replyText}
              onChangeText={setReplyText}
              multiline
              allowFontScaling={false}
              placeholder="Reply"
            />
            <View style={styles.row}>
              {panelButton('Reply', reply, !replyText.trim())}
              {button('Delete comment', deleteThread)}
              {button('Close', () => {
                setThread(null);
                setMenu(null);
              })}
            </View>
          </View>
        );
      }
      case 'link':
        return linkForm ? (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {`Link “${textOf(linkForm.para).slice(linkForm.start, linkForm.end).slice(0, 60)}” to:`}
            </Text>
            <TextInput
              style={styles.nameInput}
              value={linkForm.url}
              onChangeText={url => setLinkForm(f => (f ? {...f, url} : f))}
              autoFocus
              autoCapitalize="none"
              autoCorrect={false}
              allowFontScaling={false}
              placeholder="https://… or a DOI (10.1037/…)"
              returnKeyType="done"
              onSubmitEditing={applyLink}
            />
            <View style={styles.row}>
              {panelButton('Link', applyLink)}
              {button('Cancel', () => {
                setLinkForm(null);
                setMenu(null);
              })}
            </View>
          </View>
        ) : null;
      case 'count': {
        const all = countWords(blocks.filter((b): b is ParagraphBlock => b.type === 'p').map(paragraphText));
        const sel = !typing && selection ? countWords([selectedText()]) : null;
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {`Document: ${all.words.toLocaleString()} words, ${all.chars.toLocaleString()} characters (no spaces)`}
            </Text>
            <Text allowFontScaling={false} style={[styles.menuText, styles.countSel]}>
              {`Body text: ${words.body.toLocaleString()} words`}
            </Text>
            <Text allowFontScaling={false} style={styles.presetSummary}>
              {'Body text leaves out a title page and the reference list (and anything after it), as most assignment limits do.'}
            </Text>
            <Text allowFontScaling={false} style={[styles.menuText, styles.findLabel]}>
              {'Target (body text)'}
            </Text>
            <View style={styles.row}>
              <TextInput
                style={styles.nameInput}
                value={targetText}
                onChangeText={t => setTargetText(t.replace(/[^0-9]/g, ''))}
                keyboardType="number-pad"
                placeholder="e.g. 2500"
                allowFontScaling={false}
                onSubmitEditing={() => saveTarget(targetText)}
              />
            </View>
            {target ? (
              <Text allowFontScaling={false} style={styles.presetSummary}>
                {`${words.body >= target ? 'Reached: ' : ''}${targetLabel(words.body, target)} (${Math.round((words.body / target) * 100)}%). It shows on the bottom line.`}
              </Text>
            ) : null}
            <View style={styles.row}>
              {panelButton('Set target', () => saveTarget(targetText), targetText === '')}
              {target ? panelButton('Remove target', () => saveTarget('')) : null}
            </View>
            {sel ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.countSel]}>
                {`Selection: ${sel.words.toLocaleString()} words, ${sel.chars.toLocaleString()} characters`}
              </Text>
            ) : null}
            <View style={styles.row}>{button('Close', () => setMenu(null))}</View>
          </View>
        );
      }
      case 'view':
        return (
          <>
            <Text allowFontScaling={false} style={[styles.menuText, styles.folderPath]}>
              {`Text size: ${Math.round(textScale * 100)}%`}
            </Text>
            {item('Larger text  A+', () => setScale(scaleAt + 1))}
            {item('Smaller text  A−', () => setScale(scaleAt - 1))}
            {item('Normal size (100%)', () => setScale(SCALES.indexOf(1)))}
            {item('Pages…', () => setShowPages(true))}
            {item('Go to page…', openGoTo)}
            {item('Word count…', () => setMenu('count'))}
            {item('Go to start', () => goTo({block: 0, offset: 0}))}
            {item('Go to end', goToEnd)}
            {item(`${spellOn ? '✓ ' : ''}Mark misspellings`, () => setSpellOn(v => !v))}
            {hasNotes ? item(showMargin ? 'Fold the margin panel' : 'Open the margin panel', () => setShowMargin(v => !v)) : null}
          </>
        );
      case 'folder':
        return (
          <View>
            <Text allowFontScaling={false} style={[styles.menuText, styles.folderPath]} numberOfLines={2}>
              {shortPath(browse?.path ?? newFolder)}
            </Text>
            {browse && browse.path !== STORAGE
              ? item('↑  Up one level', () => browseTo(browse.path.slice(0, browse.path.lastIndexOf('/')) || STORAGE))
              : null}
            {browse?.folders?.map(f => item(`▸  ${f}`, () => browseTo(`${browse.path}/${f}`)))}
            {browse?.error ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.menuMuted, styles.folderPath]}>
                {'This folder can\'t be listed here.'}
              </Text>
            ) : null}
            {browse?.error ? item('Choose a file in the folder you want…', folderFromFile, styles.menuAction) : null}
            {item('Use this folder', () => {
              if (browse) {
                setNewFolder(browse.path);
              }
              setBrowse(null);
              setMenu('name');
            }, styles.menuAction)}
          </View>
        );
      case 'name':
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {template ? `New from ${template.slice(template.lastIndexOf('/') + 1)}: name` : 'Name for the new document'}
            </Text>
            <TextInput
              style={styles.nameInput}
              value={naming ?? ''}
              onChangeText={setNaming}
              autoFocus
              selectTextOnFocus
              allowFontScaling={false}
              returnKeyType="done"
              onSubmitEditing={createNew}
            />
            <View style={styles.row}>
              <Text allowFontScaling={false} style={[styles.menuText, styles.flex]} numberOfLines={2}>
                {`In ${shortPath(newFolder)}`}
              </Text>
              <Pressable onPress={once('change-folder', () => browseTo(newFolder))} style={styles.button}>
                <Text allowFontScaling={false} style={styles.buttonText}>
                  {'Change…'}
                </Text>
              </Pressable>
            </View>
            <View style={styles.row}>
              {button('Create', createNew)}
              {button('Cancel', () => {
                setNaming(null);
                setTemplate(null);
              })}
            </View>
          </View>
        );
      case 'edit':
        return (
          <>
            {item('Cut', () => copy(true))}
            {item('Copy', () => copy(false))}
            {item('Paste', pasteFromMenu)}
            {item('Delete', deleteFromMenu)}
            {item('Find & replace…', () => setMenu('find'))}
            {item('Spelling…', openSpelling)}
            {item('Link…', startLink)}
            {item('Cite from Zotero…', openCite)}
            {item('Insert quote…', openQuotes)}
            {item('Comment…', startComment)}
            {inkOk ? item('Handwritten note…', startNote) : null}
            {item('Insert picture…', startPicture)}
            {hasChanges ? item('Accept all changes', () => {
              setMenu(null);
              review(-1, '*', true);
            }) : null}
            {hasChanges ? item('Reject all changes', () => {
              setMenu(null);
              review(-1, '*', false);
            }) : null}
            {item('Remove link', removeLink)}
            {item('Hide keyboard', done)}
          </>
        );
      case 'style':
        return (
          <>
            {item('Body text', () => style('normal', 'body text'))}
            {item('Heading 1', () => style('heading1', 'heading 1'), styles.menuH1)}
            {item('Heading 2', () => style('heading2', 'heading 2'), styles.menuH2)}
            {item('Heading 3', () => style('heading3', 'heading 3'), styles.menuH3)}
            {item('Quote', () => style('quote', 'quote'), styles.menuQuote)}
            {item('Title', () => style('title', 'title'), styles.menuTitle)}
          </>
        );
      case 'hl': {
        const hc = current().hc;
        return (
          <>
            {HIGHLIGHTS.map(([value, name]) => item(`${hc === value ? '✓ ' : ''}${name}`, () => highlight(value, `highlight ${name.toLowerCase()}`)))}
            {item('None (remove highlight)', () => highlight(null, 'no highlight'))}
            <Text allowFontScaling={false} style={[styles.presetSummary, styles.folderPath]}>
              {'The color shows in Word; on the Supernote highlights are gray.'}
            </Text>
          </>
        );
      }
      case 'picture': {
        const labelHelp =
          citeStyle === 'apa'
            ? 'APA 7: “Figure N” in bold and the title in italics, above the picture.'
            : citeStyle === 'mla'
            ? 'MLA 9: “Fig. N.” and the title below the picture.'
            : 'Chicago: “Figure N.” and the title below the picture.';
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {'Insert picture'}
            </Text>
            <Text allowFontScaling={false} style={styles.presetSummary}>
              {'It goes after the paragraph you tapped, as wide as it is (never wider than the text).'}
            </Text>
            <View style={styles.row}>{panelButton(pic.path ? 'Choose another…' : 'Choose picture…', choosePicture)}</View>
            {pic.note ? (
              <Text allowFontScaling={false} style={styles.presetSummary}>
                {pic.note}
              </Text>
            ) : null}
            <Pressable onPress={once('fig-label', () => setPic(p => ({...p, label: !p.label})))} style={[styles.chip, styles.findCase, pic.label ? styles.chipOn : null]}>
              <Text allowFontScaling={false} style={[styles.chipText, pic.label ? styles.chipTextOn : null]}>
                {pic.label ? '✓ Figure label and title' : 'Figure label and title'}
              </Text>
            </Pressable>
            {pic.label ? (
              <>
                <TextInput
                  style={styles.nameInput}
                  value={pic.title}
                  onChangeText={title => setPic(p => ({...p, title}))}
                  placeholder="Figure title"
                  allowFontScaling={false}
                />
                <Text allowFontScaling={false} style={styles.presetSummary}>
                  {`${labelHelp} Numbered in order. The format follows the style chosen in Cite from Zotero (now ${CITE_STYLES.find(x => x.id === citeStyle)?.name ?? ''}).`}
                </Text>
              </>
            ) : null}
            <View style={styles.row}>
              {panelButton('Insert', insertPicture, !pic.path)}
              {button('Cancel', () => setMenu(null))}
            </View>
          </View>
        );
      }
      case 'goto':
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {map ? `Go to page (1–${map.pages.length}${map.done ? '' : '+'})` : 'Go to page'}
            </Text>
            <View style={styles.row}>
              <TextInput
                style={styles.nameInput}
                value={goToText}
                onChangeText={t => setGoToText(t.replace(/[^0-9]/g, ''))}
                keyboardType="number-pad"
                autoFocus
                allowFontScaling={false}
                returnKeyType="go"
                onSubmitEditing={goToPage}
              />
            </View>
            {goToNote ? (
              <Text allowFontScaling={false} style={styles.presetSummary}>
                {goToNote}
              </Text>
            ) : null}
            <View style={styles.row}>
              {panelButton('Go', goToPage, goToText === '')}
              {button('Close', () => setMenu(null))}
            </View>
          </View>
        );
      case 'find': {
        const n = find.query ? findAll().length : 0;
        return (
          <View style={styles.nameForm}>
            <Text allowFontScaling={false} style={styles.menuText}>
              {'Find'}
            </Text>
            <TextInput
              style={styles.nameInput}
              value={find.query}
              onChangeText={query => setFind(f => ({...f, query, at: -1}))}
              autoFocus
              autoCorrect={false}
              allowFontScaling={false}
              returnKeyType="search"
              onSubmitEditing={() => showMatch(find.at + 1)}
            />
            <Text allowFontScaling={false} style={[styles.menuText, styles.findLabel]}>
              {'Replace with'}
            </Text>
            <TextInput
              style={styles.nameInput}
              value={find.replace}
              onChangeText={replace => setFind(f => ({...f, replace}))}
              autoCorrect={false}
              allowFontScaling={false}
            />
            <Pressable onPress={once('match-case', () => setFind(f => ({...f, matchCase: !f.matchCase, at: -1})))} style={[styles.chip, styles.findCase, find.matchCase ? styles.chipOn : null]}>
              <Text allowFontScaling={false} style={[styles.chipText, find.matchCase ? styles.chipTextOn : null]}>
                {find.matchCase ? '✓ Match case' : 'Match case'}
              </Text>
            </Pressable>
            <Text allowFontScaling={false} style={styles.presetSummary}>
              {find.query ? (n === 0 ? 'Not found' : find.at >= 0 ? `${find.at + 1} of ${n}` : `${n} found`) : ' '}
            </Text>
            <View style={styles.row}>
              {panelButton('‹ Prev', () => showMatch(find.at - 1), !find.query)}
              {panelButton('Next ›', () => showMatch(find.at + 1), !find.query)}
            </View>
            <View style={styles.row}>
              {panelButton('Replace', replaceOne, !find.query)}
              {panelButton('Replace all', replaceAll, !find.query || n === 0)}
              {button('Close', () => setMenu(null))}
            </View>
          </View>
        );
      }
      case 'list':
        return (
          <>
            {item('1.  Numbered', () => listTool('number'))}
            {item('•  Bulleted', () => listTool('bullet'))}
            {item('Not a list', () => listTool('none'))}
          </>
        );
      case 'font': {
        const now = current();
        const cur = now.font;
        return (
          <>
            {item(`${now.s ? '✓ ' : ''}Strikethrough`, () => {
              setMenu(null);
              format('s', 'strikethrough');
            })}
            {item(`${(selection && !typing ? now.sup : typingScript() === 'sup') ? '✓ ' : ''}Superscript  x²`, () => scriptTool('sup'))}
            {item(`${(selection && !typing ? now.sub : typingScript() === 'sub') ? '✓ ' : ''}Subscript  x₂`, () => scriptTool('sub'))}
            <View style={styles.menuDivider} />
            {fontChoices.map(f =>
              item(
                `${f === cur ? '✓ ' : ''}${f}${fonts.has(f) ? '' : shownFont(f, fonts) ? ` (shown as ${shownFont(f, fonts)})` : ' (not on this Supernote)'}`,
                () => applyRunStyle({font: f}, `font ${f}`),
                shownFont(f, fonts) ? {fontFamily: shownFont(f, fonts)} : styles.menuMuted,
              ),
            )}
            {item('Add font files…', addFont, styles.menuAction)}
          </>
        );
      }
      case 'size': {
        const sz = current().size;
        return (
          <>
            {SIZES.map(pt => item(`${sz === pt * 2 ? '✓ ' : ''}${pt}`, () => applyRunStyle({size: pt * 2}, `size ${pt}`)))}
          </>
        );
      }
      default:
        return null;
    }
  };

  return (
    <View style={styles.root}>
      {doc && !readOnly ? (
        // Invisible: it only receives the keystrokes, which show in the page itself. Always
        // mounted, so selecting with the pen or pressing a button never drops the keyboard.
        <TextInput
          ref={keys}
          style={styles.hiddenInput}
          value={input}
          onChangeText={onKeysText}
          onKeyPress={e => e.nativeEvent.key === 'Backspace' && backspace()}
          blurOnSubmit={false}
          onSubmitEditing={onEnter}
          autoCapitalize="none"
          autoCorrect={false}
          spellCheck={false}
          caretHidden
        />
      ) : null}
      <View style={styles.header}>
        {menuButton('File', 'file')}
        <Text allowFontScaling={false} style={styles.title} numberOfLines={1}>
          {doc ? `${dirty || pending.length > 0 ? '• ' : ''}${doc.name.replace(/\.docx$/i, '')}` : 'SNScribe'}
        </Text>
        {doc ? button(contents ? 'Back to page' : 'Contents', () => setContents(c => !c), headings.length === 0) : null}
        {doc ? menuButton('View', 'view') : null}
        {doc ? button('◀', previous, atStart) : null}
        {doc ? (
          <Pressable onPress={once('pages', () => setShowPages(p => !p))} style={styles.pageCount}>
            <Text allowFontScaling={false} style={styles.page}>
              {pageNumber !== null && map ? `${pageNumber} of ${map.pages.length}${map.done ? '' : '…'}` : `${progress}%`}
            </Text>
          </Pressable>
        ) : null}
        {doc ? button('▶', next, atEnd || brk.kind === 'pending') : null}
      </View>
      <View style={styles.tools}>
        {doc && !readOnly ? (
          <ScrollView
            horizontal
            // The hidden keyboard field is usually focused; without this, a tap here first
            // only dismisses the keyboard and never reaches the button (CT: B needed two taps).
            keyboardShouldPersistTaps="always"
            showsHorizontalScrollIndicator={false}
            onScroll={e => {
              toolScroll.current = e.nativeEvent.contentOffset.x;
            }}
            scrollEventThrottle={100}
            contentContainerStyle={styles.toolsInner}>
            {button('↶', undo, edits.cursor === 0 && pending.length === 0)}
            {button('↷', redo, edits.cursor === edits.steps.length)}
            <View style={styles.divider} />
            {button('B', () => format('b', 'bold'), false, styles.square)}
            {button('I', () => format('i', 'italic'), false, styles.square)}
            {button('U', () => format('u', 'underline'), false, styles.square)}
            {menuButton('HL', 'hl')}
            <View style={styles.divider} />
            {menuButton('Edit', 'edit')}
            {menuButton('Style', 'style')}
            {menuButton('Para', 'para')}
            {menuButton('List', 'list')}
            {menuButton('Font', 'font')}
            {menuButton('Size', 'size')}
            <View style={styles.spacer} />
            {button('Save', save, !dirty && pending.length === 0)}
          </ScrollView>
        ) : null}
      </View>
      <View
        style={styles.pageArea}
        onLayout={e => {
          log(`page area ${Math.round(e.nativeEvent.layout.width)}x${Math.round(e.nativeEvent.layout.height)}`);
          setPageH(Math.floor(e.nativeEvent.layout.height) - PAD * 2);
          setPageW(Math.floor(e.nativeEvent.layout.width) - PAD * 2);
        }}>
        {doc && counting && counting.key === mapKey ? (
          <PageCounter
            key={counting.key}
            blocks={counting.blocks}
            seed={counting.seed}
            width={readW}
            pageH={pageH}
            fonts={fonts}
            scale={textScale}
            onPages={(pages, isDone) => {
              const c = counting;
              setPageMap(prev =>
                !isDone && prev && prev.done && prev.layout === c.layout ? prev : {key: c.key, layout: c.layout, blocks: c.blocks, pages, done: isDone},
              );
              if (isDone) {
                logPages(c.blocks, pages);
              }
            }}
          />
        ) : null}
        {!doc ? (
          <View style={styles.empty}>
            <Text allowFontScaling={false} style={styles.emptyText}>
              {'Open a Word document (.docx) or make a new one from the File menu.'}
            </Text>
          </View>
        ) : showPages ? (
          <ScrollView style={styles.contents} keyboardShouldPersistTaps="always">
            <View style={styles.pagesHead}>
              <Text allowFontScaling={false} style={[styles.panelTitle, styles.flex]}>
                {map ? `${map.pages.length}${map.done ? '' : '+'} pages${map.done ? '' : ' (still counting)'}` : 'Counting pages…'}
              </Text>
              <Pressable onPress={once('notes-only', () => setNotesOnly(n => !n))} style={[styles.chip, styles.pagesChip, notesOnly ? styles.chipOn : null]}>
                <Text allowFontScaling={false} style={[styles.chipText, notesOnly ? styles.chipTextOn : null]}>
                  {notesOnly ? '✓ Only pages with notes' : 'Only pages with notes'}
                </Text>
              </Pressable>
              {button('Start', () => goTo({block: 0, offset: 0}))}
              {button('End', goToEnd)}
              {button('Back to page', () => setShowPages(false))}
            </View>
            {notesOnly && pageInfos.every(pv => !pageMarks(pv)) ? (
              <Text allowFontScaling={false} style={[styles.menuText, styles.pagesNone]}>
                {'No page has comments, handwritten notes or tracked changes.'}
              </Text>
            ) : null}
            <View style={styles.pagesGrid}>
              {(map?.pages ?? []).map((pg, i) => {
                const pv = pageInfos[i];
                const marks = pv ? pageMarks(pv) : '';
                if (!pv || (notesOnly && !marks)) {
                  return null;
                }
                const here = pageNumber === i + 1;
                return (
                  <Pressable key={i} onPress={once(`page:${i}`, () => goTo(pg.anchor))} style={[styles.pageCard, here ? styles.pageCardHere : null]}>
                    <Text allowFontScaling={false} style={styles.pageCardNumber}>
                      {`${i + 1}${here ? '  ·  you are here' : ''}`}
                    </Text>
                    {pv.heading ? (
                      <Text allowFontScaling={false} style={styles.pageCardHeading} numberOfLines={1}>
                        {pv.heading}
                      </Text>
                    ) : null}
                    {marks ? (
                      <Text allowFontScaling={false} style={styles.pageCardMarks} numberOfLines={1}>
                        {marks}
                      </Text>
                    ) : null}
                    <Text allowFontScaling={false} style={styles.pageCardText} numberOfLines={marks ? 2 : 3}>
                      {pv.text}
                    </Text>
                    <View style={styles.pageCardTools}>
                      {cardButton(`newpage:${i}`, 'New page after', `Tap again to add a blank page after page ${i + 1}.`, () => newPageAfter(i))}
                      {cardButton(`selpage:${i}`, 'Select page', `Tap again to select the text of page ${i + 1}.`, () => selectPage(i))}
                      {i > 0 && pageBreakAt(i)
                        ? cardButton(`unbreak:${i}`, 'Remove page break', `Tap again to remove the page break before page ${i + 1}.`, () => removePageBreak(i))
                        : null}
                    </View>
                  </Pressable>
                );
              })}
            </View>
          </ScrollView>
        ) : contents ? (
          <ScrollView style={styles.contents} keyboardShouldPersistTaps="always">
            {headings.map(h => (
              <Pressable key={h.block} onPress={() => jump(h.block)} style={styles.contentsRow}>
                <Text allowFontScaling={false} style={[styles.contentsText, {marginLeft: Math.max(0, h.level - 1) * 24}]} numberOfLines={2}>
                  {h.text}
                </Text>
              </Pressable>
            ))}
          </ScrollView>
        ) : pageH > 0 ? (
          <View style={[styles.viewport, {height: pageH, width: textW}]}>
            <View key={pageKey} style={[styles.column, {top: -anchor.offset}]}>
              {window.map((b, i) => (
                <BlockView
                  key={anchor.block + i}
                  block={b}
                  selection={selectedIn(b)}
                  spell={spellFor(b)}
                  fonts={fonts}
                  scale={textScale}
                  width={textW}
                  onFrame={onFrame(i)}
                  onLines={onLines(i)}
                  onTextFrame={onTextFrame(i)}
                  textRef={textRefFor(i)}
                />
              ))}
              {caretAt && !typing && caretBox?.key === pageKey ? (
                <View pointerEvents="none" style={[styles.caret, {left: caretBox.x - 1, top: caretBox.top, height: caretBox.height}]} />
              ) : null}
            </View>
            <View style={[styles.mask, {top: visible}]} />
          </View>
        ) : null}
        {doc && !showPages && !contents && pageH > 0 ? (
          // Owns the pen, so nothing inks and no text handles the touch itself. It reaches
          // into the page's side margins: a pen set down just left of a line's first letter
          // starts at that letter instead of missing the page (CT).
          <TouchLayer
            style={[styles.touchLayer, {top: PAD, height: pageH, width: PAD + textW + (marginOn ? MARGIN_GAP : PAD)}]}
            onStartShouldSetResponder={() => {
              // A tap on the page while a menu is open only closes the menu (not the
              // restore question, which needs an answer).
              if (menu) {
                if (menu !== 'recover') {
                  setMenu(null);
                }
                return false;
              }
              return !readOnly;
            }}
            onMoveShouldSetResponder={() => !readOnly}
            onResponderTerminationRequest={() => false}
            onResponderGrant={e => {
              penFrom.current = point(e);
              penTo.current = penFrom.current;
            }}
            onResponderMove={e => {
              penTo.current = point(e);
            }}
            onResponderRelease={release}
            onResponderTerminate={release}
          />
        ) : null}
        {doc && marginOn && pageH > 0 ? (showMargin ? marginColumn() : marginStrip()) : null}
        {pad && pageH > 0 ? (
          <View style={[styles.pad, {width: pageW + PAD * 2, height: pageH + PAD * 2}]}>
            <View style={styles.padHead}>
              <Text allowFontScaling={false} style={[styles.menuText, styles.flex]} numberOfLines={1}>
                {pad.quote ? `Note on “${pad.quote.slice(0, 60)}”` : 'Note'}
              </Text>
              {button('Clear', () => DocxInk?.clear())}
              {button('Keep note', saveNote)}
              {button('Cancel', () => setPad(null))}
            </View>
            <View style={styles.padBody}>
            {/* Colours on the left, the writing box on the right (CT). */}
              <View style={styles.padSide}>
                <Text allowFontScaling={false} style={styles.paraLabel}>
                  {'Ink color in Word'}
                </Text>
                <View style={styles.chips}>
                  {INK_COLORS.map(([name, hex]) => (
                    <Pressable
                      key={hex}
                      onPress={once(`ink:${hex}`, () => setInkColor(hex))}
                      style={[styles.chip, inkColor === hex ? styles.chipOn : null]}>
                      <Text allowFontScaling={false} style={[styles.chipText, inkColor === hex ? styles.chipTextOn : null]}>
                        {inkColor === hex ? `✓ ${name}` : name}
                      </Text>
                    </Pressable>
                  ))}
                </View>
                <Text allowFontScaling={false} style={styles.panelNote}>
                  {'You write in black here (the screen has no color); the note is saved in the color you pick.'}
                </Text>
                <Text allowFontScaling={false} style={styles.panelNote}>
                  {'Write in the box. The note goes in the right margin of the Word file, beside these words, about 1 inch wide — Word shows and prints it.'}
                </Text>
              </View>
              {/* Fixed size and position: the pen engine must never see its host move. Narrow,
                  like the margin it goes into, so the writing stays legible when it shrinks. */}
              <View style={styles.padSurface}>
                <InkSurfaceView style={{width: PAD_W, height: pageH + PAD * 2 - PAD_HEAD - 28}} />
              </View>
            </View>
          </View>
        ) : null}
        {menu ? (
          <View style={[styles.menu, {left: Math.max(0, Math.min(menuX, pageW + PAD * 2 - MENU_W - 4)), maxHeight: Math.max(240, pageH)}]}>
            <ScrollView keyboardShouldPersistTaps="always">{menuBody()}</ScrollView>
          </View>
        ) : null}
      </View>
      <View style={styles.statusLine}>
        <Text allowFontScaling={false} style={styles.statusText} numberOfLines={1}>
          {status ||
            (typing
              ? typing.mode === 'insert'
                ? 'Typing. Enter starts a new paragraph.'
                : 'Type to replace the selection.'
              : caret
              ? 'Type to insert at the caret.'
              : doc && !readOnly
              ? 'Pen: tap to type · drag to select · double-tap a word. Finger: swipe to turn pages.'
              : '')}
        </Text>
        {doc && target ? (
          <Pressable onPress={once('count', () => setMenu('count'))} style={styles.statusPage}>
            <Text allowFontScaling={false} style={styles.statusPageText}>
              {`${words.body >= target ? '✓ ' : ''}${targetLabel(words.body, target)}`}
            </Text>
          </Pressable>
        ) : null}
        {doc && pageNumber !== null && map ? (
          <Pressable onPress={once('goto', openGoTo)} style={styles.statusPage}>
            <Text allowFontScaling={false} style={styles.statusPageText}>
              {`Page ${pageNumber} of ${map.pages.length}${map.done ? '' : '+'}`}
            </Text>
          </Pressable>
        ) : null}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  root: {flex: 1, backgroundColor: '#fff'},
  header: {
    height: HEADER_H,
    flexDirection: 'row',
    alignItems: 'center',
    paddingHorizontal: PAD,
    borderBottomWidth: 1,
    borderColor: '#000',
  },
  title: {flex: 1, color: '#000', fontSize: 20, fontWeight: '700', marginHorizontal: 12},
  page: {color: '#000', fontSize: 16, minWidth: 72, textAlign: 'center'},
  button: {borderWidth: 1, borderColor: '#000', borderRadius: 5, paddingVertical: 8, paddingHorizontal: 14, marginLeft: 8},
  buttonText: {color: '#000', fontSize: 16, fontWeight: '700'},
  disabled: {opacity: 0.3},
  tools: {height: BAR_H, borderBottomWidth: 1, borderColor: '#000'},
  toolsInner: {flexGrow: 1, alignItems: 'center', paddingHorizontal: PAD - 8, paddingRight: PAD},
  divider: {width: 1, height: 30, backgroundColor: '#000', marginLeft: 12, marginRight: 4},
  spacer: {flex: 1, minWidth: 16},
  square: {minWidth: 44, alignItems: 'center', paddingHorizontal: 10},
  small: {paddingVertical: 4, paddingHorizontal: 10},
  buttonOpen: {backgroundColor: '#000'},
  buttonOpenText: {color: '#fff'},
  row: {flexDirection: 'row', alignItems: 'center', marginTop: 10},
  flex: {flex: 1},
  folderPath: {paddingHorizontal: 16, paddingVertical: 10, fontWeight: '700'},
  pageCount: {paddingHorizontal: 4},
  paraMenu: {padding: 12},
  paraLabel: {color: '#000', fontSize: 15, fontWeight: '700', marginTop: 8, marginBottom: 4},
  chips: {flexDirection: 'row', flexWrap: 'wrap'},
  chip: {borderWidth: 1, borderColor: '#000', borderRadius: 5, paddingVertical: 8, paddingHorizontal: 12, marginRight: 8, marginBottom: 8},
  chipOn: {backgroundColor: '#000'},
  chipText: {color: '#000', fontSize: 16},
  chipTextOn: {color: '#fff'},
  countSel: {marginTop: 8},
  presetSummary: {color: '#333', fontSize: 14, marginTop: 4},
  pageCardTools: {flexDirection: 'row', flexWrap: 'wrap', marginTop: 'auto'},
  pageCardButton: {borderWidth: 1, borderColor: '#000', borderRadius: 4, paddingVertical: 5, paddingHorizontal: 6, marginRight: 6, marginTop: 6},
  pageCardButtonArmed: {backgroundColor: '#000'},
  pageCardButtonTextArmed: {color: '#fff'},
  pageCardButtonText: {color: '#000', fontSize: 12, fontWeight: '700'},
  pagesHead: {flexDirection: 'row', alignItems: 'center', paddingVertical: 8, borderBottomWidth: 1, borderColor: '#000'},
  pagesGrid: {flexDirection: 'row', flexWrap: 'wrap', paddingTop: 8},
  pageCard: {width: 220, minHeight: 210, borderWidth: 1, borderColor: '#000', padding: 10, marginRight: 12, marginBottom: 12},
  pageCardHere: {borderWidth: 3},
  pageCardNumber: {color: '#000', fontSize: 18, fontWeight: '700'},
  pageCardHeading: {color: '#000', fontSize: 15, fontWeight: '700', marginTop: 4},
  pageCardText: {color: '#333', fontSize: 14, marginTop: 4},
  pageCardMarks: {color: '#000', fontSize: 14, fontWeight: '700', marginTop: 4, textDecorationLine: 'underline'},
  pagesChip: {marginBottom: 0, marginLeft: 8},
  pagesNone: {padding: 16},
  touchLayer: {position: 'absolute', left: 0},
  menu: {position: 'absolute', top: 0, width: MENU_W, backgroundColor: '#fff', borderWidth: 2, borderColor: '#000'},
  menuItem: {paddingVertical: 14, paddingHorizontal: 16, borderBottomWidth: 1, borderColor: '#bbb'},
  menuText: {color: '#000', fontSize: 19},
  menuH1: {fontSize: 24, fontWeight: '700'},
  menuH2: {fontSize: 21, fontWeight: '700'},
  menuH3: {fontSize: 19, fontWeight: '700'},
  menuQuote: {fontStyle: 'italic', paddingLeft: 32},
  menuDivider: {height: 3, backgroundColor: '#000', marginVertical: 4},
  findLabel: {marginTop: 10},
  italicText: {fontStyle: 'italic'},
  spellWord: {fontWeight: '700', textDecorationLine: 'underline', color: '#000'},
  citeMessage: {marginTop: 8, fontWeight: '700'},
  findCase: {alignSelf: 'flex-start', marginTop: 10},
  menuTitle: {fontSize: 26},
  menuMuted: {color: '#555'},
  menuAction: {fontWeight: '700'},
  nameForm: {padding: 16},
  statusLine: {height: 40, flexDirection: 'row', alignItems: 'center', paddingHorizontal: PAD, borderTopWidth: 1, borderColor: '#000'},
  actions: {flex: 1},
  actionsInner: {alignItems: 'center'},
  editButtons: {flexDirection: 'row', marginLeft: 8},
  typing: {flex: 1, flexDirection: 'row', alignItems: 'center'},
  hiddenInput: {position: 'absolute', left: 0, top: 0, width: 1, height: 1, opacity: 0},
  caretRow: {flex: 1, flexDirection: 'row', alignItems: 'center'},
  nameLabel: {color: '#000', fontSize: 16, marginRight: 8},
  nameInput: {flex: 1, height: 44, borderWidth: 1, borderColor: '#000', paddingHorizontal: 10, fontSize: 18, color: '#000'},
  panelHead: {flexDirection: 'row', alignItems: 'center', paddingVertical: 8, borderBottomWidth: 1, borderColor: '#000'},
  panelTitle: {flex: 1, color: '#000', fontSize: 18, fontWeight: '700'},
  panelNote: {color: '#333', fontSize: 15, marginTop: 16},
  hfChips: {marginTop: 8},
  presetName: {paddingHorizontal: 12, paddingTop: 8, gap: 10},
  caretHint: {flex: 1, color: '#000', fontSize: 15},
  caret: {position: 'absolute', width: 3, backgroundColor: '#000'},
  statusText: {flex: 1, color: '#000', fontSize: 15},
  statusPage: {marginLeft: 12, paddingVertical: 6, paddingHorizontal: 10, borderWidth: 1, borderColor: '#000', borderRadius: 5},
  statusPageText: {color: '#000', fontSize: 15, fontWeight: '700'},
  pageArea: {flex: 1, padding: PAD},
  viewport: {overflow: 'hidden'},
  column: {position: 'absolute', left: 0, right: 0},
  margin: {position: 'absolute', width: MARGIN_W, borderLeftWidth: 1, borderColor: '#999', paddingLeft: 8},
  card: {
    position: 'absolute',
    left: 8,
    right: 0,
    height: CARD_H,
    borderWidth: 1,
    borderColor: '#000',
    borderRadius: 6,
    flexDirection: 'row',
    alignItems: 'center',
    paddingLeft: 8,
    backgroundColor: '#fff',
  },
  cardText: {flex: 1},
  inkCard: {position: 'absolute', left: 8, right: 0, borderWidth: 1, borderColor: '#000', borderRadius: 6, padding: 3, backgroundColor: '#fff', alignItems: 'center'},
  pad: {position: 'absolute', left: 0, top: 0, backgroundColor: '#fff'},
  padBody: {flexDirection: 'row', padding: 12},
  padSurface: {borderWidth: 2, borderColor: '#000'},
  padSide: {flex: 1, marginRight: 16},
  strip: {position: 'absolute', width: MARGIN_STRIP, borderLeftWidth: 1, borderColor: '#999', alignItems: 'center', paddingTop: 8},
  stripArrow: {color: '#000', fontSize: 30, fontWeight: '700'},
  stripCount: {color: '#000', fontSize: 16, marginTop: 8, borderWidth: 1, borderColor: '#000', borderRadius: 12, minWidth: 24, textAlign: 'center', paddingHorizontal: 4},
  foldButton: {position: 'absolute', left: 8, right: 0, height: 40, borderWidth: 1, borderColor: '#000', borderRadius: 6, justifyContent: 'center', alignItems: 'center', backgroundColor: '#fff'},
  foldText: {color: '#000', fontSize: 16},
  padHead: {height: PAD_HEAD, flexDirection: 'row', alignItems: 'center', paddingHorizontal: 12, borderBottomWidth: 1, borderColor: '#000'},
  commentCard: {borderStyle: 'dashed', paddingRight: 8},
  commentInput: {height: 110, textAlignVertical: 'top', paddingTop: 8, marginVertical: 8},
  threadReply: {marginTop: 10, paddingLeft: 14, borderLeftWidth: 2, borderColor: '#000'},
  cardHead: {color: '#000', fontSize: 14, fontWeight: '700'},
  cardBody: {color: '#333', fontSize: 15, marginTop: 2},
  cardInserted: {textDecorationLine: 'underline'},
  cardDeleted: {textDecorationLine: 'line-through'},
  cardButton: {width: 44, height: CARD_H - 2, alignItems: 'center', justifyContent: 'center', borderLeftWidth: 1, borderColor: '#000'},
  cardButtonText: {color: '#000', fontSize: 24},
  cardMore: {position: 'absolute', left: 8, right: 0, color: '#333', fontSize: 13},
  mask: {position: 'absolute', left: 0, right: 0, bottom: 0, backgroundColor: '#fff'},
  empty: {flex: 1, alignItems: 'center', justifyContent: 'center'},
  emptyText: {color: '#000', fontSize: 20},
  contents: {flex: 1},
  contentsRow: {paddingVertical: 12, borderBottomWidth: 1, borderColor: '#ccc'},
  contentsText: {color: '#000', fontSize: 19},
});
