// Parsing of SharePoint $batch responses and classification of each operation.
// Pure functions: no DOM, no chrome.* APIs.

import {
  CRLF,
  normalizeLineEndings,
  splitMultipart,
  splitHeadersAndBody,
  parseHeaderLines,
  boundaryFromContentType,
} from './batch.js';

// "Already exists" errors make re-runs safe, so they are reported as warnings.
// English and German SharePoint messages are covered.
const ALREADY_EXISTS = [
  /already exists/i,
  /duplicate field name/i,
  /bereits vorhanden/i,
  /existiert bereits/i,
  /doppelter feldname/i,
];

export const DEFAULT_RETRY_AFTER_MS = 10_000;

/** Pull code and message out of an OData error body (verbose or nometadata). */
export function extractError(bodyText) {
  if (!bodyText || !bodyText.trim()) return { code: null, message: null };
  let json;
  try {
    json = JSON.parse(bodyText);
  } catch {
    return { code: null, message: bodyText.trim().slice(0, 500) };
  }
  const err = json['odata.error'] || json.error || (json.d && json.d.error);
  if (!err) return { code: null, message: null };
  const msg = err.message;
  const message = typeof msg === 'string' ? msg : msg && (msg.value || msg.Value) || null;
  return { code: err.code || null, message };
}

function parseHttpResponse(text) {
  const lines = text.split(CRLF);
  let i = 0;
  while (i < lines.length && lines[i].trim() === '') i++;
  const statusLine = (lines[i] || '').trim();
  const m = /^HTTP\/\d(?:\.\d)?\s+(\d{3})\s*(.*)$/.exec(statusLine);
  const headerLines = [];
  for (i++; i < lines.length && lines[i].trim() !== ''; i++) headerLines.push(lines[i]);
  const body = lines.slice(i + 1).join(CRLF).replace(/\s+$/, '');
  return {
    status: m ? Number(m[1]) : null,
    statusText: m ? m[2] : statusLine,
    headers: parseHeaderLines(headerLines),
    body,
  };
}

/**
 * Parse a multipart $batch response into a flat list of HTTP responses,
 * in order. Nested changeset responses are flattened.
 * @param {string} text response body
 * @param {string} [contentType] response Content-Type header (boundary source)
 */
export function parseBatchResponse(text, contentType) {
  const body = normalizeLineEndings(text);
  let boundary = boundaryFromContentType(contentType);
  if (!boundary) {
    const first = body.split(CRLF).find((l) => l.trim() !== '') || '';
    if (!first.startsWith('--')) throw new Error('Response is not a multipart batch response.');
    boundary = first.trim().slice(2);
  }
  const out = [];
  const walk = (content, b) => {
    for (const part of splitMultipart(content, b).parts) {
      if (part.trim() === '') continue;
      const { headers, rest } = splitHeadersAndBody(part);
      const ct = headers['content-type'] || '';
      if (/^multipart\/mixed/i.test(ct)) {
        const inner = boundaryFromContentType(ct);
        if (inner) walk(rest, inner);
        continue;
      }
      const res = parseHttpResponse(rest);
      const { code, message } = res.status >= 400 ? extractError(res.body) : { code: null, message: null };
      out.push({ ...res, errorCode: code, errorMessage: message });
    }
  };
  walk(body, boundary);
  return out;
}

export function isAlreadyExists(message) {
  return !!message && ALREADY_EXISTS.some((re) => re.test(message));
}

export function isThrottled(status) {
  return status === 429 || status === 503;
}

/** success | warning (already exists) | throttled | failed */
export function classify(response) {
  if (!response || response.status == null) return 'failed';
  if (response.status >= 200 && response.status < 300) return 'success';
  if (isThrottled(response.status)) return 'throttled';
  if (isAlreadyExists(response.errorMessage)) return 'warning';
  return 'failed';
}

/** Retry-After as seconds or HTTP date; falls back to 10 s. */
export function retryAfterMs(value, now = Date.now()) {
  if (value == null || String(value).trim() === '') return DEFAULT_RETRY_AFTER_MS;
  const v = String(value).trim();
  if (/^\d+(\.\d+)?$/.test(v)) return Math.round(Number(v) * 1000);
  const date = Date.parse(v);
  if (!Number.isNaN(date)) return Math.max(0, date - now);
  return DEFAULT_RETRY_AFTER_MS;
}

/**
 * Pair parsed operations with responses by position.
 * Operations without a response are reported as failed.
 */
export function correlate(operations, responses) {
  return operations.map((op, i) => {
    const res = responses[i];
    const result = {
      index: op.index,
      line: op.line,
      method: op.method,
      url: op.url,
      status: res ? res.status : null,
      statusText: res ? res.statusText : '',
      errorCode: res ? res.errorCode : null,
      errorMessage: res ? res.errorMessage : 'No response returned for this operation.',
      retryAfter: res ? res.headers['retry-after'] || null : null,
    };
    result.outcome = classify(res ? result : null);
    return result;
  });
}

export function summarize(results) {
  const s = { total: results.length, succeeded: 0, warnings: 0, failed: 0, throttled: 0 };
  for (const r of results) {
    if (r.outcome === 'success') s.succeeded++;
    else if (r.outcome === 'warning') s.warnings++;
    else if (r.outcome === 'throttled') s.throttled++;
    else s.failed++;
  }
  return s;
}
