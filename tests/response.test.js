import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseBatchResponse,
  extractError,
  classify,
  correlate,
  summarize,
  retryAfterMs,
  isAlreadyExists,
  DEFAULT_RETRY_AFTER_MS,
} from '../extension/lib/response.js';
import { parseBatch } from '../extension/lib/batch.js';
import { readFixture } from './helpers.js';

const CT = 'multipart/mixed; boundary=batchresponse_5e1c9a2f-7b3d-4c11-9a0e-2f6d8b1c4a77';

test('fixture response parses into five operations (LF and CRLF)', () => {
  const lf = readFixture('response-mixed.txt');
  for (const text of [lf, lf.replace(/\n/g, '\r\n')]) {
    const res = parseBatchResponse(text, CT);
    assert.equal(res.length, 5);
    assert.deepEqual(res.map((r) => r.status), [201, 400, 429, 400, 204]);
    assert.equal(res[0].statusText, 'Created');
    assert.equal(JSON.parse(res[0].body).Title, 'crRequests');
    assert.equal(res[2].headers['retry-after'], '7');
    assert.equal(res[4].body, '');
  }
});

test('boundary falls back to the first line when the header is missing', () => {
  const res = parseBatchResponse(readFixture('response-mixed.txt'), 'multipart/mixed');
  assert.equal(res.length, 5);
  assert.throws(() => parseBatchResponse('{"odata.error":{}}', 'application/json'), /not a multipart/);
});

test('error messages are extracted from nometadata and verbose bodies', () => {
  const res = parseBatchResponse(readFixture('response-mixed.txt'), CT);
  assert.equal(res[0].errorMessage, null);
  assert.equal(res[1].errorMessage, 'A duplicate field name "crStatus" was found.');
  assert.equal(res[1].errorCode, '-2130575245, System.ArgumentException');
  assert.match(res[3].errorMessage, /Field type Lookupp is not installed properly/);
  assert.deepEqual(extractError('not json at all'), { code: null, message: 'not json at all' });
  assert.deepEqual(extractError(''), { code: null, message: null });
  assert.equal(extractError('{"error":{"code":"1","message":"plain"}}').message, 'plain');
});

test('classification: success, already-exists warning, throttled, failure', () => {
  const res = parseBatchResponse(readFixture('response-mixed.txt'), CT);
  assert.deepEqual(res.map(classify), ['success', 'warning', 'throttled', 'failed', 'success']);
  assert.equal(classify({ status: 503, errorMessage: null }), 'throttled');
  assert.equal(classify(null), 'failed');
});

test('already-exists detection covers list and field duplicates in EN and DE', () => {
  assert.ok(isAlreadyExists('A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title.'));
  assert.ok(isAlreadyExists('A duplicate field name "crStatus" was found.'));
  assert.ok(isAlreadyExists('Eine Liste, eine Umfrage, ein Diskussionsforum oder eine Dokumentbibliothek mit dem angegebenen Titel ist auf dieser Website bereits vorhanden.'));
  assert.ok(isAlreadyExists('Es wurde ein doppelter Feldname "crStatus" gefunden.'));
  assert.ok(!isAlreadyExists('Access denied.'));
  assert.ok(!isAlreadyExists(null));
});

test('correlate pairs request operations with responses by position', () => {
  const ops = parseBatch(readFixture('request-mixed.txt')).operations;
  const results = correlate(ops, parseBatchResponse(readFixture('response-mixed.txt'), CT));
  assert.equal(results.length, 5);
  assert.equal(results[0].method, 'POST');
  assert.equal(results[4].method, 'MERGE');
  assert.equal(results[2].retryAfter, '7');
  assert.deepEqual(summarize(results), { total: 5, succeeded: 2, warnings: 1, failed: 1, throttled: 1 });

  const short = correlate(ops, parseBatchResponse(readFixture('response-mixed.txt'), CT).slice(0, 3));
  assert.equal(short[4].outcome, 'failed');
  assert.match(short[4].errorMessage, /No response/);
});

test('Retry-After: seconds, HTTP date, default 10 s', () => {
  assert.equal(retryAfterMs('7'), 7000);
  assert.equal(retryAfterMs(undefined), DEFAULT_RETRY_AFTER_MS);
  assert.equal(retryAfterMs(''), DEFAULT_RETRY_AFTER_MS);
  assert.equal(retryAfterMs('garbage'), DEFAULT_RETRY_AFTER_MS);
  const now = Date.parse('2026-10-08T12:00:00Z');
  assert.equal(retryAfterMs('Thu, 08 Oct 2026 12:00:30 GMT', now), 30_000);
  assert.equal(retryAfterMs('Thu, 08 Oct 2026 11:00:00 GMT', now), 0);
});
