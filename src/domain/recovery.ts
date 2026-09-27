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
 * Unsaved edits of one document: the steps made since the file on disk was last written
 * (by DOCX or anything else), and that file's fingerprint. Restoring is only offered when
 * the file still has that fingerprint — the steps apply to exactly that content.
 */
export type RecoveryRecord = {path: string; stamp: string; time: number; steps: Op[][]};

export function recoveryFor(
  path: string,
  stamp: string | null,
  steps: Op[][],
  cursor: number,
  savedCursor: number,
  pending: Op[],
): RecoveryRecord | null {
  // Undoing past the last save and editing on leaves nothing on disk to replay onto.
  if (!stamp || savedCursor < 0 || savedCursor > cursor) {
    return null;
  }
  const since = steps.slice(savedCursor, cursor);
  if (pending.length > 0) {
    since.push(pending);
  }
  return since.length > 0 ? {path, stamp, time: Date.now(), steps: since} : null;
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
