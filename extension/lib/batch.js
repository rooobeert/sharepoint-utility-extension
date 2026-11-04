// Parsing and rebuilding of SharePoint REST $batch request bodies
// (multipart/mixed with an outer batch boundary and optional changesets).
// Pure functions: no DOM, no chrome.* APIs, so they run in node --test as well.

export const CRLF = '\r\n';

// SharePoint Online rejects batches with more than 100 operations.
export const MAX_OPERATIONS_PER_BATCH = 100;

/** Convert any mix of CRLF / LF / CR to CRLF and drop a leading BOM. */
export function normalizeLineEndings(text) {
  return String(text).replace(/^﻿/, '').replace(/\r\n|\r|\n/g, CRLF);
}

/** The batch boundary is the first non-empty line, which must start with "--". */
export function detectBoundary(text) {
  const lines = normalizeLineEndings(text).split(CRLF);
  const first = lines.find((l) => l.trim() !== '');
  if (!first || !first.startsWith('--') || first.trim().length <= 2) {
    throw new Error('Batch body must start with a boundary line such as "--batch_xyz".');
  }
  return first.trim().slice(2).replace(/--$/, '');
}

/** Parse "Name: value" lines into an object with lower-case keys. */
export function parseHeaderLines(lines) {
  const headers = {};
  for (const line of lines) {
    const i = line.indexOf(':');
    if (i > 0) headers[line.slice(0, i).trim().toLowerCase()] = line.slice(i + 1).trim();
  }
  return headers;
}

/** Read the boundary parameter from a Content-Type header value. */
export function boundaryFromContentType(contentType) {
  const m = /boundary\s*=\s*("([^"]+)"|[^;\s]+)/i.exec(contentType || '');
  return m ? (m[2] || m[1]) : null;
}

/**
 * Split a CRLF multipart body on the given boundary.
 * Returns the raw text of each part (without delimiter lines), the 0-based
 * line index in `text` where each part starts, and whether a closing
 * delimiter "--boundary--" was seen.
 */
export function splitMultipart(text, boundary) {
  const lines = text.split(CRLF);
  const open = '--' + boundary;
  const close = open + '--';
  const parts = [];
  const starts = [];
  let current = null;
  let closed = false;
  for (const [n, line] of lines.entries()) {
    const t = line.trimEnd();
    if (t === close) {
      if (current) parts.push(current.join(CRLF));
      current = null;
      closed = true;
      break;
    }
    if (t === open) {
      if (current) parts.push(current.join(CRLF));
      current = [];
      starts.push(n + 1);
      continue;
    }
    if (current) current.push(line);
  }
  if (current) parts.push(current.join(CRLF));
  return { parts, starts, closed, lineCount: lines.length };
}

/** Split a part into its header block and the rest (after the first blank line). */
export function splitHeadersAndBody(partText) {
  const lines = partText.split(CRLF);
  let i = 0;
  while (i < lines.length && lines[i].trim() === '') i++; // tolerate leading blank lines
  const headerLines = [];
  for (; i < lines.length && lines[i].trim() !== ''; i++) headerLines.push(lines[i]);
  return { headers: parseHeaderLines(headerLines), rest: lines.slice(i + 1).join(CRLF), restStart: i + 1 };
}

/** Parse the embedded HTTP request of an application/http part. */
function parseHttpRequest(text) {
  const lines = text.split(CRLF);
  let i = 0;
  while (i < lines.length && lines[i].trim() === '') i++;
  const requestLine = (lines[i] || '').trim();
  const m = /^([A-Z]+)\s+(\S+)(?:\s+HTTP\/\d(?:\.\d)?)?$/.exec(requestLine);
  const headerLines = [];
  for (i++; i < lines.length && lines[i].trim() !== ''; i++) headerLines.push(lines[i]);
  const body = lines.slice(i + 1).join(CRLF).replace(/\s+$/, '');
  return {
    method: m ? m[1] : null,
    url: m ? m[2] : null,
    requestLine,
    requestIndex: lines.findIndex((l) => l.trim() !== ''),
    headers: parseHeaderLines(headerLines),
    body,
  };
}

/**
 * Prepare a pasted batch body for sending: CRLF line endings, no leading
 * blank lines, one trailing CRLF. The content itself is not changed; a
 * missing closing boundary is reported by parseBatch as an error.
 */
export function prepareBatchBody(text) {
  const boundary = detectBoundary(text);
  const body = normalizeLineEndings(text).replace(/^(\s*\r\n)+/, '').replace(/\s+$/, '');
  return { body: body + CRLF, boundary };
}

const WRITE_METHODS = new Set(['POST', 'MERGE', 'PATCH', 'PUT', 'DELETE']);

/**
 * Parse a batch body into its operations.
 * Each operation keeps its raw MIME part text so a subset can be re-sent
 * byte-for-byte (used for retrying throttled operations), and the 1-based
 * line number of its request line in the text as pasted.
 * `errors` are structural problems that make the body unsendable (they block
 * Run); `warnings` never block.
 */
export function parseBatch(text) {
  const { body, boundary } = prepareBatchBody(text);
  // Line numbers refer to the text as the user sees it, so parse it without
  // the leading-blank-line trimming that prepareBatchBody does for sending.
  const source = normalizeLineEndings(text);
  const errors = [];
  const warnings = [];
  const operations = [];
  const outer = splitMultipart(source, boundary);
  if (!outer.closed) {
    errors.push(`The batch is never closed: the closing line "--${boundary}--" is missing (the text ends at line ${outer.lineCount}).`);
  }

  // `at` is the 0-based line index of the part's first line in `source`.
  const addOperation = (raw, at, changeset) => {
    const { headers: mime, rest, restStart } = splitHeadersAndBody(raw);
    const req = parseHttpRequest(rest);
    const line = at + restStart + Math.max(0, req.requestIndex) + 1;
    const n = operations.length + 1;
    if (!req.method || !req.url) {
      errors.push(`Operation ${n} (line ${line}): could not read the request line "${req.requestLine}".`);
    } else if (!changeset && WRITE_METHODS.has(req.method)) {
      errors.push(
        `Operation ${n} (line ${line}): ${req.method} is outside a changeset. SharePoint only accepts GET outside a changeset and rejects the whole batch with "invalid HTTP method '${req.method}' … for a query operation".`
      );
    }
    operations.push({
      index: operations.length,
      line,
      method: req.method,
      url: req.url,
      headers: req.headers,
      body: req.body,
      mimeHeaders: mime,
      changeset,
      raw,
    });
  };

  outer.parts.forEach((part, k) => {
    if (part.trim() === '') return;
    const at = outer.starts[k];
    const { headers, rest, restStart } = splitHeadersAndBody(part);
    const ct = headers['content-type'] || '';
    if (/^multipart\/mixed/i.test(ct)) {
      const cs = boundaryFromContentType(ct);
      if (!cs) {
        errors.push(`Line ${at + 1}: a changeset has no boundary in its Content-Type.`);
        return;
      }
      const inner = splitMultipart(rest, cs);
      const innerAt = at + restStart;
      if (!inner.closed) {
        errors.push(`Changeset "${cs}" (started at line ${at}) is never closed: the closing line "--${cs}--" is missing.`);
      }
      if (!inner.starts.length) {
        errors.push(`Changeset "${cs}" (started at line ${at}) contains no "--${cs}" line, so SharePoint sees no operations in it.`);
      }
      inner.parts.forEach((p, j) => {
        if (p.trim() !== '') addOperation(p, innerAt + inner.starts[j], cs);
      });
    } else {
      addOperation(part, at, null);
    }
  });
  if (operations.length === 0) errors.push('No operations found in this batch.');
  return { boundary, body, operations, errors, warnings };
}

/**
 * Rebuild a batch body from a subset of parsed operations, keeping the
 * original boundaries and part text. Consecutive operations of the same
 * changeset are grouped into one changeset again.
 */
export function buildBatch(boundary, operations) {
  let out = '';
  let i = 0;
  while (i < operations.length) {
    const cs = operations[i].changeset;
    if (cs) {
      out += `--${boundary}${CRLF}Content-Type: multipart/mixed; boundary=${cs}${CRLF}${CRLF}`;
      while (i < operations.length && operations[i].changeset === cs) {
        out += `--${cs}${CRLF}${operations[i].raw}${CRLF}`;
        i++;
      }
      out += `--${cs}--${CRLF}`;
    } else {
      out += `--${boundary}${CRLF}${operations[i].raw}${CRLF}`;
      i++;
    }
  }
  return out + `--${boundary}--${CRLF}`;
}

/** Trim whitespace and trailing slashes. */
export function normalizeSiteUrl(url) {
  return String(url || '').trim().replace(/\/+$/, '');
}

/** Checks that only warn: batches above SharePoint's operation limit. */
export function analyseBatch(parsed) {
  const warnings = [];
  if (parsed.operations.length > MAX_OPERATIONS_PER_BATCH) {
    warnings.push(
      `${parsed.operations.length} operations: SharePoint Online accepts at most ${MAX_OPERATIONS_PER_BATCH} per batch. Split this stage.`
    );
  }
  return warnings;
}

/** Site part of an absolute operation URL (everything before "/_api/"), or null. */
export function siteOfOperationUrl(url) {
  const m = /^(https:\/\/[^/]+(?:\/[^?#]*?)?)\/_api\//i.exec(url || '');
  return m ? m[1] : null;
}

/**
 * Work out the one site all stages target, from the absolute operation URLs.
 * stages: [{ name, parsed }] with parsed from parseBatch.
 * Returns { site, sites, relativeOnly } where `sites` lists every distinct
 * site with the stages and operation numbers that use it. `site` is null when
 * the stages disagree or contain no absolute URLs (relativeOnly = true).
 */
export function targetSiteFromStages(stages) {
  const bySite = new Map();
  for (const [s, { name, parsed }] of stages.entries()) {
    for (const op of parsed.operations) {
      const site = siteOfOperationUrl(op.url);
      if (!site) continue;
      const key = site.toLowerCase();
      if (!bySite.has(key)) bySite.set(key, { site, uses: [] });
      const entry = bySite.get(key);
      let use = entry.uses.find((u) => u.stage === s);
      if (!use) entry.uses.push((use = { stage: s, name, operations: [] }));
      use.operations.push(op.index + 1);
    }
  }
  const sites = [...bySite.values()];
  return { site: sites.length === 1 ? sites[0].site : null, sites, relativeOnly: sites.length === 0 };
}

/** Derive a site URL from a page URL: https://host/sites/x or /teams/x, else the root. */
export function siteUrlFromPageUrl(pageUrl) {
  let u;
  try {
    u = new URL(pageUrl);
  } catch {
    return null;
  }
  const m = /^\/(sites|teams|portals)\/[^/]+/i.exec(u.pathname);
  return u.origin + (m ? m[0] : '');
}

/** Only https://<tenant>.sharepoint.com URLs are allowed as targets. */
export function isSharePointUrl(url) {
  try {
    const u = new URL(url);
    return u.protocol === 'https:' && /^[a-z0-9-]+\.sharepoint\.com$/i.test(u.hostname);
  } catch {
    return false;
  }
}

/** True when `url` is an absolute URL on exactly `origin` (scheme, host and port). */
export function isOnOrigin(url, origin) {
  try {
    return new URL(url).origin === origin;
  } catch {
    return false;
  }
}
