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

async function get(account: ZoteroAccount, path: string): Promise<unknown> {
  let res: Response;
  try {
    res = await fetch(`${API}/users/${encodeURIComponent(account.userId.trim())}${path}`, {
      headers: {'Zotero-API-Key': account.apiKey.trim(), 'Zotero-API-Version': '3'},
    });
  } catch {
    throw new Error('Could not reach Zotero. Is Wi-Fi on?');
  }
  if (!res.ok) {
    throw new Error(failure(res.status));
  }
  return res.json();
}

/** Checks the account can read the library. */
export async function testZotero(account: ZoteroAccount): Promise<void> {
  await get(account, '/items/top?limit=1&format=json');
}

/** Top-level items matching `query` (title, creators, year), formatted in `csl`. */
export async function searchZotero(account: ZoteroAccount, query: string, csl: string): Promise<Source[]> {
  const q = encodeURIComponent(query.trim());
  const items = (await get(
    account,
    `/items/top?q=${q}&qmode=titleCreatorYear&format=json&include=data,bib,citation&style=${encodeURIComponent(csl)}&limit=25`,
  )) as Item[];
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
