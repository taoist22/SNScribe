// A DOI's details from doi.org (content negotiation: CSL JSON), for sources not in Zotero.

import {detailsFromCsl, type Details} from '../domain/reference';

export async function lookUpDoi(doi: string, log: (line: string) => void = () => {}): Promise<Details> {
  const url = `https://doi.org/${encodeURIComponent(doi.trim()).replace(/%2F/g, '/')}`;
  const started = Date.now();
  let res: Response;
  try {
    res = await Promise.race([
      fetch(url, {headers: {Accept: 'application/vnd.citationstyles.csl+json'}}),
      new Promise<never>((_, reject) => setTimeout(() => reject(new Error('timeout')), 20000)),
    ]);
  } catch (error) {
    const why = error instanceof Error ? error.message : String(error);
    log(`doi GET ${url} failed after ${Date.now() - started} ms: ${why}`);
    throw new Error(why === 'timeout' ? 'doi.org did not answer in 20 seconds.' : `Could not reach doi.org (${why}).`);
  }
  log(`doi GET ${url} → ${res.status} in ${Date.now() - started} ms`);
  if (res.status === 404) {
    throw new Error(`doi.org does not know ${doi}. Check the DOI.`);
  }
  if (!res.ok) {
    throw new Error(`doi.org answered ${res.status}.`);
  }
  return detailsFromCsl((await res.json()) as Record<string, unknown>);
}
