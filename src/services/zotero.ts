// The Zotero Web API (api.zotero.org, v3): search the user's synced library and get each
// item's in-text citation and reference entry formatted by Zotero in a CSL style. Needs a
// read-only API key and the numeric user ID (zotero.org → Settings → Security).

import type {Source} from '../domain/citations';

export type ZoteroAccount = {userId: string; apiKey: string};

const API = 'https://api.zotero.org';

type Item = {
  key: string;
  data?: {itemType?: string; title?: string; date?: string};
  meta?: {creatorSummary?: string; parsedDate?: string};
  bib?: string;
  citation?: string;
};

/** A readable reason for a failed request. */
function failure(status: number): string {
  if (status === 403) {
    return 'Zotero refused the key or user ID. Check both in Zotero settings.';
  }
  if (status === 404) {
    return 'Zotero has no library for that user ID.';
  }
  if (status === 429 || status === 503) {
    return 'Zotero is busy. Try again in a minute.';
  }
  return `Zotero answered ${status}.`;
}

/** Where requests are reported (the log): never the key. */
let report: (line: string) => void = () => {};
export function setZoteroLog(fn: (line: string) => void): void {
  report = fn;
}

const TIMEOUT_MS = 20000;

async function get(account: ZoteroAccount, path: string): Promise<unknown> {
  const url = `${API}/users/${encodeURIComponent(account.userId.trim())}${path}`;
  const started = Date.now();
  let res: Response;
  const timer = new Promise<never>((_, reject) => setTimeout(() => reject(new Error('timeout')), TIMEOUT_MS));
  try {
    res = await Promise.race([
      fetch(url, {headers: {'Zotero-API-Key': account.apiKey.trim(), 'Zotero-API-Version': '3'}}),
      timer,
    ]);
  } catch (error) {
    const why = error instanceof Error ? error.message : String(error);
    report(`zotero GET ${url} failed after ${Date.now() - started} ms: ${why}`);
    throw new Error(why === 'timeout' ? 'Zotero did not answer in 20 seconds. Check Wi-Fi and try again.' : `Could not reach Zotero (${why}). Is Wi-Fi on?`);
  }
  report(`zotero GET ${url} → ${res.status} in ${Date.now() - started} ms (key ${account.apiKey.trim().length} chars)`);
  if (!res.ok) {
    let detail = '';
    try {
      detail = (await res.text()).slice(0, 120);
    } catch {
      // No body.
    }
    if (detail) {
      report(`zotero said: ${detail}`);
    }
    throw new Error(failure(res.status));
  }
  return res.json();
}

/**
 * A Zotero key and user ID found in text the user keeps (a .txt): the key is 24 letters and
 * digits; the user ID is a run of digits on its own ("userID: 1234567" or a bare number).
 */
export function credentialsIn(text: string): {apiKey?: string; userId?: string} {
  // Word by word (no lookbehind: the device's JavaScript engine may not support it).
  const words = text.split(/[^A-Za-z0-9]+/).filter(Boolean);
  const apiKey = words.find(w => w.length === 24);
  const userId = words.find(w => /^\d{3,12}$/.test(w));
  return {...(apiKey ? {apiKey} : {}), ...(userId ? {userId} : {})};
}

/** Checks the account can read the library. */
export async function testZotero(account: ZoteroAccount): Promise<void> {
  await get(account, '/items/top?limit=1&format=json');
}

/** One item of the library by its key, formatted in `csl`. */
export async function zoteroItem(account: ZoteroAccount, key: string, csl: string): Promise<Source> {
  const it = (await get(account, `/items/${encodeURIComponent(key)}?format=json&include=data,bib,citation&style=${encodeURIComponent(csl)}`)) as Item;
  if (!it?.bib || !it.citation) {
    throw new Error('Zotero sent no citation for that item.');
  }
  return {
    key: it.key,
    title: it.data?.title ?? '',
    authors: it.meta?.creatorSummary ?? '',
    year: (it.meta?.parsedDate ?? it.data?.date ?? '').slice(0, 4),
    citation: it.citation,
    bibHtml: it.bib,
  };
}

/** Top-level items matching `query` (title, creators, year), formatted in `csl`. */
export async function searchZotero(account: ZoteroAccount, query: string, csl: string): Promise<Source[]> {
  const q = encodeURIComponent(query.trim());
  const items = (await get(
    account,
    `/items/top?q=${q}&qmode=titleCreatorYear&format=json&include=data,bib,citation&style=${encodeURIComponent(csl)}&limit=25`,
  )) as Item[];
  report(`zotero: ${Array.isArray(items) ? items.length : 'no list'} item(s) back; with a citation: ${Array.isArray(items) ? items.filter(it => it.bib && it.citation).length : 0}`);
  return items
    .filter(it => it.data && !['note', 'attachment', 'annotation'].includes(it.data.itemType ?? '') && it.bib && it.citation)
    .map(it => ({
      key: it.key,
      title: it.data?.title ?? '',
      authors: it.meta?.creatorSummary ?? '',
      year: (it.meta?.parsedDate ?? it.data?.date ?? '').slice(0, 4),
      citation: it.citation!,
      bibHtml: it.bib!,
    }));
}
