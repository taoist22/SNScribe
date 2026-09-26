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
  /** Saves a verified copy with `ops` applied; `dest` '' = new <name>-edited.docx beside the original. */
  save(src: string, ops: object[], dest: string): Promise<{dest: string; name: string; ms: number; changed: string[]}>;
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
};

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
