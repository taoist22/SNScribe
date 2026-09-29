// Recovery records and recent documents: small pure helpers, stored as JSON in DOCX's
// private storage (DocxModule.store/load).

import type {Op} from './edits';

/** A short, stable key for a document path (names private files). */
export function docKey(path: string): string {
  let h = 5381;
  for (let i = 0; i < path.length; i++) {
    h = ((h << 5) + h + path.charCodeAt(i)) >>> 0;
  }
  const name = path.split('/').pop() ?? 'doc';
  return `${name.replace(/[^A-Za-z0-9]/g, '').slice(0, 24)}-${h.toString(16)}`;
}

/**
 * Unsaved edits of one document, and the fingerprint of the file on disk when they were
 * recorded. Restoring is only offered when the file still has that fingerprint.
 *
 * Usually `steps` are the edits since the file was last written, replayed onto the file.
 * After undoing past the last save and editing on, nothing on disk matches any step: then
 * `base` names a private copy of the document as it was opened, and `steps` are every edit
 * since (the editor's own history starts there).
 */
export type RecoveryRecord = {path: string; stamp: string; time: number; steps: Op[][]; base?: string};

/** What the journal should hold now: a record, nothing ('clean': all saved), or the base copy still to be made. */
export type Journal = {kind: 'clean'} | {kind: 'record'; record: RecoveryRecord} | {kind: 'needsBase'; steps: Op[][]};

export function journalFor(
  path: string,
  stamp: string | null,
  steps: Op[][],
  cursor: number,
  savedCursor: number,
  pending: Op[],
  base: string | null,
): Journal | null {
  if (!stamp) {
    return null; // nothing to fingerprint against: leave the journal as it is
  }
  if (cursor === savedCursor && pending.length === 0) {
    return {kind: 'clean'};
  }
  if (savedCursor >= 0 && savedCursor <= cursor) {
    const since = steps.slice(savedCursor, cursor);
    if (pending.length > 0) {
      since.push(pending);
    }
    return {kind: 'record', record: {path, stamp, time: Date.now(), steps: since}};
  }
  // Undone past the last save and edited on: every step since opening, on the opened copy.
  const all = steps.slice(0, cursor);
  if (pending.length > 0) {
    all.push(pending);
  }
  return base ? {kind: 'record', record: {path, stamp, time: Date.now(), steps: all, base}} : {kind: 'needsBase', steps: all};
}

/** The record for the common case (edits since the last save), or null. Kept for tests and callers that need only that. */
export function recoveryFor(path: string, stamp: string | null, steps: Op[][], cursor: number, savedCursor: number, pending: Op[]): RecoveryRecord | null {
  const j = journalFor(path, stamp, steps, cursor, savedCursor, pending, null);
  return j?.kind === 'record' ? j.record : null;
}

export function parseRecovery(json: string | null, path: string, stamp: string | null): RecoveryRecord | null {
  if (!json || !stamp) {
    return null;
  }
  try {
    const r = JSON.parse(json) as RecoveryRecord;
    return r.path === path && r.stamp === stamp && Array.isArray(r.steps) && r.steps.length > 0 ? r : null;
  } catch {
    return null;
  }
}

export type RecentDoc = {path: string; name: string; time: number};

/** The recent list with `path` moved to the front; ten at most. */
export function touchRecent(list: RecentDoc[], path: string, name: string): RecentDoc[] {
  return [{path, name, time: Date.now()}, ...list.filter(r => r.path !== path)].slice(0, 10);
}
