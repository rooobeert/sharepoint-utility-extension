// Stage runner. Orchestration runs in the side panel; every HTTP call is made
// inside the SharePoint tab (see page* functions below, injected with
// chrome.scripting.executeScript) so requests are same-origin and carry the
// user's session cookies. The form digest is fetched and used inside the page
// and never returned to the extension.

import { parseBatch, buildBatch, analyseBatch, normalizeSiteUrl } from './batch.js';
import {
  parseBatchResponse,
  correlate,
  summarize,
  isThrottled,
  retryAfterMs,
  extractError,
} from './response.js';
import {
  findPlaceholders,
  resolvePlaceholders,
  applyPlaceholders,
  describePlaceholder,
  findCreator,
  findUrlPlaceholders,
  applyUrlPlaceholders,
  urlPlaceholderValues,
} from './placeholders.js';

export const MAX_RETRIES = 3;

// ---------------------------------------------------------------------------
// Injected functions. They run in the SharePoint page, so they must not
// reference anything outside their own body.
// ---------------------------------------------------------------------------

/** MAIN world: read the web and site collection URLs SharePoint publishes for the page. */
export function pageReadContext() {
  const info = window._spPageContextInfo;
  return {
    webAbsoluteUrl: (info && info.webAbsoluteUrl) || null,
    siteAbsoluteUrl: (info && info.siteAbsoluteUrl) || null,
    href: location.href,
  };
}

/**
 * ISOLATED world: perform one request. With `digestFrom` set, a fresh
 * FormDigestValue is fetched from {digestFrom}/_api/contextinfo first and sent
 * as X-RequestDigest. Only status, type, two headers and the body text come back.
 * `redirect` is passed to fetch (default 'follow').
 */
export async function pageRequest(req) {
  const pick = (res) => ({
    'content-type': res.headers.get('content-type'),
    'retry-after': res.headers.get('retry-after'),
  });
  try {
    const headers = Object.assign({}, req.headers);
    if (req.digestFrom) {
      const ci = await fetch(req.digestFrom + '/_api/contextinfo', {
        method: 'POST',
        headers: { Accept: 'application/json;odata=nometadata' },
        credentials: 'same-origin',
        cache: 'no-store',
      });
      if (!ci.ok) {
        return { phase: 'contextinfo', status: ci.status, statusText: ci.statusText, headers: pick(ci), text: await ci.text() };
      }
      headers['X-RequestDigest'] = (await ci.json()).FormDigestValue;
    }
    const res = await fetch(req.url, {
      method: req.method,
      headers,
      body: req.body,
      credentials: 'same-origin',
      cache: 'no-store',
      redirect: req.redirect || 'follow',
    });
    // With redirect: 'manual' a redirect comes back as type 'opaqueredirect' (status 0, no body).
    return { phase: 'request', status: res.status, statusText: res.statusText, type: res.type, headers: pick(res), text: await res.text() };
  } catch (e) {
    return { phase: 'network', status: 0, statusText: String((e && e.message) || e), headers: {}, text: '' };
  }
}

// ---------------------------------------------------------------------------

export class StageError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.details = details;
  }
}

const defaultSleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * SharePoint rejects a malformed batch as a whole and doesn't say which
 * operation caused it. Point at the likely operations by line number.
 */
export function locateBatchError(message, operations) {
  const hints = [];
  const m = /HTTP method '(\w+)'/i.exec(message || '');
  if (m) {
    const method = m[1].toUpperCase();
    const ops = operations.filter((op) => op.method === method);
    const list = ops.slice(0, 5).map((op) => `operation ${op.index + 1} at line ${op.line}`).join(', ');
    if (ops.length) hints.push(`${method} operations in this stage: ${list}${ops.length > 5 ? ` and ${ops.length - 5} more` : ''}.`);
    if (/query operation/i.test(message)) {
      hints.push(
        `SharePoint read a ${method} as outside a changeset, so it probably did not recognise the changeset. Check the lines before line ${ops[0] ? ops[0].line : '?'}: the batch must start with "--<batch boundary>", then "Content-Type: multipart/mixed; boundary=<changeset boundary>", an empty line, and "--<changeset boundary>" before each operation.`
      );
    }
  }
  const lineRef = /line (\d+)/i.exec(message || '');
  if (!hints.length && !lineRef && operations.length) {
    hints.push(`SharePoint rejected the whole batch, so no operation ran. The stage has ${operations.length} operation(s), lines ${operations[0].line}–${operations[operations.length - 1].line}.`);
  }
  return hints;
}

function httpError(res, what) {
  if (res.phase === 'network') return `${what} failed: ${res.statusText}`;
  const { message } = extractError(res.text);
  const where = res.phase === 'contextinfo' ? 'Getting the form digest (/_api/contextinfo)' : what;
  return `${where} returned HTTP ${res.status} ${res.statusText || ''}`.trim() + (message ? `: ${message}` : '');
}

/**
 * @param {object} deps
 * @param {(req: object) => Promise<object>} deps.request  runs pageRequest in the tab
 * @param {(ms: number) => Promise<void>} [deps.sleep]
 * @param {(event: object) => void} [deps.onEvent]  progress events for the UI
 */
export function createRunner({ request, sleep = defaultSleep, onEvent = () => {} }) {
  const fetchJson = (site) => async (path) => {
    const res = await request({
      url: site + path,
      method: 'GET',
      headers: { Accept: 'application/json;odata=nometadata' },
    });
    if (res.phase === 'network') throw new Error(res.statusText);
    let json = null;
    try {
      json = JSON.parse(res.text);
    } catch {
      /* non-JSON error page */
    }
    return { status: res.status, json, message: extractError(res.text).message };
  };

  /** Look up all placeholders of a body. Returns { resolved, listGuids }. */
  async function resolve(site, text) {
    const found = findPlaceholders(text);
    if (!found.length) return { resolved: [], listGuids: null };
    return resolvePlaceholders(found, fetchJson(site));
  }

  // Verified web/site URLs, read once per runner (for the URL placeholders).
  const siteInfoCache = new Map();
  function siteInfo(site) {
    if (!siteInfoCache.has(site)) siteInfoCache.set(site, verifySite(site));
    return siteInfoCache.get(site);
  }

  /** Replace {{WebUrl}} etc. Returns { text, values } (values only for used names). */
  async function withUrlPlaceholders(site, text) {
    const names = findUrlPlaceholders(text);
    if (!names.length) return { text, values: {} };
    const all = urlPlaceholderValues(await siteInfo(site));
    const values = Object.fromEntries(names.map((n) => [n, all[n]]));
    return { text: applyUrlPlaceholders(text, all), values };
  }

  /**
   * Ask SharePoint what the target URL really is: web title and URL, and the
   * site collection URL. Throws StageError if the site can't be read.
   */
  async function verifySite(siteUrl) {
    const site = normalizeSiteUrl(siteUrl);
    const get = fetchJson(site);
    const web = await get('/_api/web?$select=Title,Url');
    if (web.status < 200 || web.status >= 300 || !web.json || !web.json.Url) {
      throw new StageError(
        `Could not read the site ${site} (HTTP ${web.status}${web.message ? ': ' + web.message : ''}). Check the URL and your access.`
      );
    }
    const sc = await get('/_api/site?$select=Url');
    if (sc.status < 200 || sc.status >= 300 || !sc.json || !sc.json.Url) {
      throw new StageError(`Could not read the site collection of ${site} (HTTP ${sc.status}).`);
    }
    return { title: web.json.Title, webUrl: normalizeSiteUrl(web.json.Url), siteCollectionUrl: normalizeSiteUrl(sc.json.Url) };
  }

  /** POST a batch; on 429/503 wait for Retry-After and try again, up to MAX_RETRIES times. */
  async function sendBatch(site, body, boundary, ctx) {
    for (let attempt = 0; ; attempt++) {
      onEvent({ type: 'send', stage: ctx.stage, attempt, operations: ctx.count });
      const res = await request({
        url: site + '/_api/$batch',
        method: 'POST',
        headers: {
          'Content-Type': `multipart/mixed; boundary=${boundary}`,
          Accept: 'application/json;odata=nometadata',
        },
        body,
        digestFrom: site,
      });
      if (!isThrottled(res.status) || attempt >= MAX_RETRIES) return res;
      const ms = retryAfterMs(res.headers && res.headers['retry-after']);
      const wait = { ms, reason: `HTTP ${res.status} on the batch request`, retry: attempt + 1, of: MAX_RETRIES };
      ctx.waits.push(wait);
      onEvent({ type: 'wait', stage: ctx.stage, ...wait });
      await sleep(ms);
    }
  }

  async function sendAndParse(site, body, boundary, ctx, operations) {
    const res = await sendBatch(site, body, boundary, ctx);
    if (res.status < 200 || res.status >= 300) {
      // Digest and network failures happen before SharePoint reads the batch.
      const read = res.phase !== 'contextinfo' && res.phase !== 'network';
      const hints = read ? locateBatchError(extractError(res.text).message, operations) : [];
      throw new StageError([httpError(res, 'The $batch request'), ...hints].join(' '));
    }
    try {
      return parseBatchResponse(res.text, res.headers && res.headers['content-type']);
    } catch (e) {
      throw new StageError(`Could not read the $batch response: ${e.message}`);
    }
  }

  /**
   * Plan the parts of a stage: an operation that uses a list which doesn't
   * exist yet, but is created by an earlier operation of the same stage,
   * starts a new part. Returns { splits: [{ index, line, ref, creator }], problems }.
   */
  function planParts(operations, placeholders) {
    const splits = new Map();
    const problems = [];
    for (const p of placeholders) {
      if (p.id || p.type !== 'list') continue;
      const first = operations.find((op) => findPlaceholders(op.raw).some((x) => x.key === p.key));
      if (!first) continue;
      const creator = findCreator(operations, p.ref, first.index);
      const later = creator ? null : findCreator(operations, p.ref);
      if (creator) {
        p.pending = { line: creator.line, operation: creator.index + 1 };
        if (!splits.has(first.index)) splits.set(first.index, { index: first.index, line: first.line, ref: p.ref, creator });
      } else if (later) {
        problems.push(`${describePlaceholder(p)} is used at line ${first.line} before it is created at line ${later.line}. Move the list creation before its first use.`);
      }
    }
    return { splits: [...splits.values()].sort((a, b) => a.index - b.index), problems };
  }

  /** Parse, analyse and resolve placeholders without sending the batch. */
  async function preview(siteUrl, stage) {
    const site = normalizeSiteUrl(siteUrl);
    const { text, values } = await withUrlPlaceholders(site, stage.body);
    const parsed = parseBatch(text);
    const warnings = [...parsed.warnings, ...analyseBatch(parsed)];
    const { resolved, listGuids } = await resolve(site, text);
    const { splits, problems } = planParts(parsed.operations, resolved);
    warnings.push(...problems);
    const placeholders = [
      ...Object.entries(values).map(([name, value]) => ({ label: `{{${name}}}`, id: value, error: null })),
      ...resolved.map((p) => ({ label: describePlaceholder(p), tokens: p.tokens, id: p.id, error: p.error, pending: p.pending || null })),
    ];
    const notes = splits.length
      ? [`Sent in ${splits.length + 1} parts: lists created in this stage are looked up before line ${splits.map((x) => x.line).join(', ')}.`]
      : [];
    return { name: stage.name, operations: parsed.operations, errors: parsed.errors, warnings, notes, placeholders, listGuids, parts: splits.length + 1 };
  }

  /** Send operations as one batch, re-sending throttled ones. Returns results in `ops` order. */
  async function sendOps(site, boundary, ops, body, ctx) {
    ctx.count = ops.length;
    const first = correlate(ops, await sendAndParse(site, body, boundary, ctx, ops));
    const byIndex = new Map(first.map((r) => [r.index, { ...r, attempts: 1 }]));
    const opByIndex = new Map(ops.map((o) => [o.index, o]));

    // Operations throttled inside the batch: re-send only those, unchanged.
    for (let round = 1; round <= MAX_RETRIES; round++) {
      const idx = [...byIndex.values()].filter((r) => r.outcome === 'throttled').map((r) => r.index);
      if (!idx.length) break;
      const ms = Math.max(...idx.map((i) => retryAfterMs(byIndex.get(i).retryAfter)));
      const wait = { ms, reason: `${idx.length} operation(s) throttled`, retry: round, of: MAX_RETRIES };
      ctx.waits.push(wait);
      onEvent({ type: 'wait', stage: ctx.stage, ...wait });
      await sleep(ms);
      const retry = idx.map((i) => opByIndex.get(i));
      ctx.count = retry.length;
      let again;
      try {
        again = correlate(retry, await sendAndParse(site, buildBatch(boundary, retry), boundary, ctx, retry));
      } catch (e) {
        // Keep the results so far; the throttled ones stay throttled.
        ctx.result.warnings.push(`Retry of throttled operations failed: ${e.message}`);
        break;
      }
      for (const r of again) byIndex.set(r.index, { ...r, attempts: round + 1 });
    }
    return ops.map((o) => {
      const r = byIndex.get(o.index);
      if (r.outcome !== 'throttled') return r;
      return { ...r, outcome: 'failed', errorMessage: `Still throttled after ${MAX_RETRIES} retries` + (r.errorMessage ? `: ${r.errorMessage}` : '.') };
    });
  }

  const notSent = (op, why) => ({
    index: op.index,
    line: op.line,
    method: op.method,
    url: op.url,
    status: null,
    statusText: '',
    outcome: 'failed',
    errorMessage: `Not sent: ${why}`,
  });

  /**
   * Run one stage. Lookups are resolved right before the operations that need
   * them: if a list is created inside this stage, the stage is sent in parts,
   * in the original order, and the lists are read again between the parts.
   * Throws StageError only if nothing could be sent.
   */
  async function runStage(siteUrl, stage, stageIndex) {
    const site = normalizeSiteUrl(siteUrl);
    const result = {
      name: stage.name,
      startedAt: new Date().toISOString(),
      placeholders: [],
      warnings: [],
      notes: [],
      waits: [],
      operations: [],
    };
    const ctx = { stage: stageIndex, waits: result.waits, count: 0, result };

    // 1. URL placeholders, then structure. Line numbers stay those of the editor.
    let text;
    let urlValues;
    try {
      ({ text, values: urlValues } = await withUrlPlaceholders(site, stage.body));
    } catch (e) {
      throw new StageError(`Could not fill in the URL placeholders, nothing was sent: ${e.message}`, { result });
    }
    const parsed = parseBatch(text);
    if (parsed.errors.length) throw new StageError(`The batch body is invalid, nothing was sent: ${parsed.errors.join(' ')}`, { result });
    result.warnings = [...parsed.warnings, ...analyseBatch(parsed)];
    const ops = parsed.operations;
    const opKeys = ops.map((op) => findPlaceholders(op.raw).map((p) => p.key));
    const all = findPlaceholders(text);
    const known = new Map(); // key -> resolved placeholder
    const failedLookups = new Map(); // key -> last unresolved placeholder
    const results = [];
    const splitLines = [];
    let pos = 0;
    let parts = 0;

    const report = () => {
      result.placeholders = [
        ...Object.entries(urlValues).map(([name, value]) => ({ label: `{{${name}}}`, id: value, error: null })),
        ...all.map((p) => {
          const hit = known.get(p.key) || failedLookups.get(p.key) || p;
          return { label: describePlaceholder(p), id: hit.id || null, error: hit.id ? null : hit.error || 'Not looked up.' };
        }),
      ];
    };

    while (pos < ops.length) {
      // 2. Resolve what the remaining operations need and isn't known yet.
      const needed = all.filter((p) => !known.has(p.key) && opKeys.slice(pos).some((keys) => keys.includes(p.key)));
      if (needed.length) {
        onEvent({ type: 'resolve', stage: stageIndex });
        const { resolved, listGuids } = await resolvePlaceholders(needed, fetchJson(site));
        if (listGuids) result.listGuids = listGuids;
        for (const p of resolved) {
          if (p.id) known.set(p.key, p);
          else failedLookups.set(p.key, p);
        }
      }

      // 3. A missing list that no earlier operation of this stage creates can
      // never be resolved: stop before sending anything more.
      const firstUse = (key) => opKeys.findIndex((keys, k) => k >= pos && keys.includes(key));
      const doomed = all
        .filter((p) => !known.has(p.key) && firstUse(p.key) !== -1)
        .map((p) => ({ p: failedLookups.get(p.key) || p, at: firstUse(p.key) }))
        .filter(({ p, at }) => p.type === 'web' || !findCreator(ops, p.ref, at))
        .sort((a, b) => a.at - b.at)[0];

      // 4. Otherwise everything up to the first operation whose list is created
      // in an earlier part.
      let end = pos;
      while (end < ops.length && opKeys[end].every((k) => known.has(k))) end++;

      if (doomed || end === pos) {
        const at = doomed ? doomed.at : pos;
        const op = ops[at];
        const missing = doomed ? [doomed.p] : opKeys[pos].filter((k) => !known.has(k)).map((k) => failedLookups.get(k));
        const why = missing.map((p) => {
          const creator = p.type === 'list' && findCreator(ops, p.ref, at);
          const made = creator && results.find((r) => r.index === creator.index);
          let text = `${describePlaceholder(p)} (${p.error || 'not found'})`;
          if (!creator) text += ' No earlier operation in this stage creates it.';
          else if (made && made.outcome === 'failed') {
            text += ` It should have been created by operation ${creator.index + 1} (line ${creator.line}), which failed: ${made.errorMessage}`;
          } else if (made) {
            text += ` Operation ${creator.index + 1} (line ${creator.line}) created a list with that title, but the reference doesn't match its URL. Use the list's title or its URL.`;
          }
          return text;
        });
        const message = `Operation ${op.index + 1} (line ${op.line}) needs ${why.join('; ')}`;
        report();
        if (!parts) throw new StageError(`Unresolved placeholders, nothing was sent: ${message}`, { result });
        result.error = `${message} The rest of the stage (from line ${ops[pos].line}) was not sent.`;
        for (const o of ops.slice(pos)) results.push(notSent(o, `a lookup needed at line ${op.line} could not be resolved.`));
        break;
      }

      // 5. Send this part. The whole stage goes out byte-for-byte; parts are
      // rebuilt from the original operation texts.
      const chunk = ops.slice(pos, end).map((op) => ({ ...op, raw: applyPlaceholders(op.raw, [...known.values()]) }));
      const whole = pos === 0 && end === ops.length;
      const body = whole ? applyPlaceholders(parsed.body, [...known.values()]) : buildBatch(parsed.boundary, chunk);
      parts++;
      if (pos > 0) splitLines.push(ops[pos].line);
      if (!whole) onEvent({ type: 'part', stage: stageIndex, part: parts, from: ops[pos].line, operations: chunk.length });
      try {
        results.push(...(await sendOps(site, parsed.boundary, chunk, body, ctx)));
      } catch (e) {
        if (!(e instanceof StageError)) throw e;
        report();
        if (parts === 1) {
          e.details.result = result;
          throw e;
        }
        result.error = `${e.message} Earlier parts were sent; this part and the rest of the stage were not run.`;
        for (const o of ops.slice(pos)) results.push(notSent(o, `SharePoint rejected the part starting at line ${ops[pos].line}.`));
        break;
      }
      pos = end;
    }

    report();
    if (parts > 1) {
      result.notes.push(`Sent in ${parts} parts: lists created in this stage were looked up again before line ${splitLines.join(', ')}.`);
    }
    result.parts = parts;
    result.operations = results;
    result.summary = summarize(results);
    result.finishedAt = new Date().toISOString();
    return result;
  }

  /**
   * Run stages in order. After a stage with real failures (or a stage error)
   * `shouldContinue(stageResult, index)` decides whether to go on.
   */
  async function runStages(siteUrl, stages, { shouldContinue = async () => false } = {}) {
    const run = { site: normalizeSiteUrl(siteUrl), startedAt: new Date().toISOString(), stages: [] };
    for (let i = 0; i < stages.length; i++) {
      onEvent({ type: 'stage-start', stage: i, name: stages[i].name });
      let stageResult;
      try {
        stageResult = { ...(await runStage(siteUrl, stages[i], i)), status: 'completed' };
      } catch (e) {
        const partial = (e.details && e.details.result) || { name: stages[i].name, operations: [], waits: [] };
        stageResult = { ...partial, status: 'error', error: e.message, finishedAt: new Date().toISOString() };
      }
      run.stages.push(stageResult);
      onEvent({ type: 'stage-done', stage: i, result: stageResult });

      const failed = stageResult.status === 'error' || stageResult.summary.failed > 0;
      if (failed && i < stages.length - 1 && !(await shouldContinue(stageResult, i))) {
        for (let j = i + 1; j < stages.length; j++) {
          run.stages.push({ name: stages[j].name, status: 'skipped', operations: [], waits: [] });
        }
        break;
      }
    }
    run.finishedAt = new Date().toISOString();
    return run;
  }

  return { verifySite, preview, runStage, runStages };
}
