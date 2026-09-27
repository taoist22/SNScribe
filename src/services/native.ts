// Typed handle on the native Docx module, and the session log.

import {NativeModules} from 'react-native';
import type {DocxDocument} from '../model/docx';

type DocxModule = {
  NATIVE_BUILD?: number;
  getConstants?: () => {NATIVE_BUILD: number};
  /** Resolves null when written, else why the write was refused. */
  log(line: string): Promise<string | null>;
  logName(): Promise<string>;
  delay(ms: number): Promise<void>;
  open(path: string): Promise<DocxDocument>;
  /** A blank document named `name` in the Document folder: its path, and the pristine copy edits apply to. */
  create(name: string): Promise<{path: string; source: string}>;
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
  ): Promise<{dest: string; name: string; ms: number; changed: string[]}>;
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
};

/** A key the native listener caught: DEL_FWD, arrows, HOME/END, or a Ctrl/Cmd letter; `text` for a paste. */
export type KeyPress = {key: string; shift: boolean; text?: string};

type DocxKeysModule = {
  attach(tag: number): Promise<string>;
  copy(text: string): Promise<boolean>;
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
