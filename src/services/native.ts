// Typed handle on the native Docx module, and the session log.

import {NativeModules} from 'react-native';
import type {DocxDocument, ParagraphBlock} from '../model/docx';

type DocxModule = {
  NATIVE_BUILD?: number;
  getConstants?: () => {NATIVE_BUILD: number};
  /** Resolves null when written, else why the write was refused. */
  log(line: string): Promise<string | null>;
  logName(): Promise<string>;
  delay(ms: number): Promise<void>;
  open(path: string): Promise<DocxDocument>;
  /** A blank document named `name` in the Document folder: its path, and the pristine copy edits apply to. */
  create(name: string, folder: string): Promise<{path: string; source: string}>;
  /** A new document in folder copied from a .docx or .dotx (made a document). */
  createFrom(templatePath: string, name: string, folder: string): Promise<{path: string; source: string}>;
  /** Folders directly inside `path`, or {error} when it can't be listed. */
  listFolders(path: string): Promise<{folders?: string[]; error?: string}>;
  /** DOCX's private storage: JSON kept under a name. */
  store(name: string, json: string): Promise<string | null>;
  load(name: string): Promise<string | null>;
  forget(name: string): Promise<boolean>;
  /** A private copy of `source` kept under `name` (a recovery record's base); resolves its path. Missing before native build 21. */
  keepBase?(source: string, name: string): Promise<string>;
  /** "size:modified" of a file, or null when it is not there. */
  fileStamp(path: string): Promise<string | null>;
  /** A private copy of the document as opened; edits apply to it. */
  snapshot(path: string, key: string): Promise<string>;
  /** Backs the file up (newest five kept); resolves the backup's path. */
  backup(path: string, key: string): Promise<string>;
  backups(key: string): Promise<Array<{path: string; time: number; bytes: number}>>;
  copyOver(from: string, dest: string): Promise<boolean>;
  /** Where a copy of `path` goes: <name>-edited.docx beside it. */
  copyName(path: string): Promise<string>;
  /** Loads what it can find of these families (plus fonts added before); resolves the families available. */
  fonts(families: string[]): Promise<string[]>;
  /** A font file the user picked: its family, loaded and remembered. */
  addFont(path: string): Promise<{family: string; style: number}>;
  /** Saves a verified copy with `ops` applied; `dest` '' = new <name>-edited.docx beside the original. */
  save(
    src: string,
    ops: object[],
    dest: string,
    /** Every paragraph's text as the screen shows it after the edits; the save is refused on any mismatch. */
    expected: string[],
    /** The document's key (recovery/backups): a save that fails while writing keeps its checked file there. */
    key: string,
  ): Promise<{dest: string; name: string; ms: number; changed: string[]}>;
  /** The misspelled ones among `words` ([] loads the dictionary). */
  spellCheck(lang: string, words: string[]): Promise<string[]>;
  /** Corrections for a misspelled word, closest first. */
  spellSuggest(lang: string, word: string): Promise<string[]>;
  /** An EPUB's own title, creators, date and publisher. */
  epubInfo(path: string): Promise<{title: string; creators: string[]; date: string; publisher: string}>;
  /** A small text file's contents (at most 16 KB). */
  readText(path: string): Promise<string>;
  /** Paragraphs `paras` (empty = all) as they will be once `ops` are applied to `src`; nothing is written. */
  preview(src: string, ops: object[], paras: number[]): Promise<{blocks: ParagraphBlock[]}>;
};

export type OffsetResult = {
  offset?: number;
  char?: number;
  line?: number;
  lineCount?: number;
  length?: number;
  viewClass: string;
  error?: string;
};

type DocxTextModule = {
  /** x, y in px relative to the Text view: the gap nearest the pen and the character under it. */
  offsetAt(tag: number, x: number, y: number): Promise<OffsetResult>;
  /** Where a caret before character `offset` goes: px relative to the Text view. */
  caretRect(tag: number, offset: number): Promise<{x?: number; top?: number; bottom?: number; error?: string}>;
  /** The offset one line up (dir -1) or down (+1) at the same x, or {outside} past the paragraph. */
  lineMove(tag: number, offset: number, dir: number): Promise<{offset?: number; outside?: boolean; error?: string}>;
  /** Whether the last touch on the page (DocxTouchLayer) was the pen or a finger. Missing before native build 14. */
  gestureTool?(): Promise<{tool: 'pen' | 'finger'; toolType: number}>;
  /** Where characters [start, end) of a Text view sit, one box per line (px relative to the view). Missing before native build 22. */
  rangeRects?(tag: number, start: number, end: number): Promise<{rects?: Array<{left: number; right: number; top: number; bottom: number}>; error?: string}>;
};

/** A key the native listener caught: DEL_FWD, arrows, HOME/END, or a Ctrl/Cmd letter; `text` for a paste. */
export type KeyPress = {key: string; shift: boolean; text?: string};

type DocxKeysModule = {
  attach(tag: number): Promise<string>;
  /** Shows the on-screen keyboard for the field (not with a hardware keyboard). Missing before native build 23. */
  showKeyboard?(tag: number): Promise<string>;
  copy(text: string): Promise<boolean>;
  /** The clipboard's text, or null when it is empty or can't be read. */
  clipboardText(): Promise<string | null>;
};

export const DocxKeys = NativeModules.DocxKeys as DocxKeysModule | undefined;

export const DocxText = NativeModules.DocxText as DocxTextModule | undefined;

export const Docx = NativeModules.Docx as DocxModule | undefined;

export function nativeBuild(): number | string {
  return Docx?.NATIVE_BUILD ?? Docx?.getConstants?.().NATIVE_BUILD ?? 'missing';
}

/** Appends to EXPORT/sn-docx-log-*.txt. Never throws; resolves the refusal reason or null. */
export async function log(...lines: string[]): Promise<string | null> {
  try {
    let refused: string | null = null;
    for (const line of lines) {
      refused = (await Docx?.log(line)) ?? refused;
    }
    return refused;
  } catch (error) {
    return errorText(error);
  }
}

export function errorText(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}
