import { test } from 'node:test';
import assert from 'node:assert/strict';
import { formatRunText, formatRunJson, shortUrl } from '../extension/lib/log.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const run = {
  site: SITE,
  startedAt: '2026-10-08T12:00:00.000Z',
  finishedAt: '2026-10-08T12:00:09.000Z',
  stages: [
    {
      name: 'lists',
      status: 'completed',
      placeholders: [{ label: '{{WebId}}', id: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee' }],
      waits: [{ ms: 7000, reason: '1 operation(s) throttled', retry: 1, of: 3 }],
      warnings: [],
      summary: { total: 3, succeeded: 1, warnings: 1, failed: 1, throttled: 0 },
      operations: [
        { index: 0, line: 8, method: 'POST', url: SITE + '/_api/web/lists', status: 201, outcome: 'success', attempts: 2 },
        { index: 1, method: 'POST', url: SITE + '/_api/web/lists', status: 500, outcome: 'warning', errorMessage: 'A list ... already exists' },
        { index: 2, method: 'POST', url: SITE + '/_api/web/lists', status: 400, outcome: 'failed', errorMessage: 'Boom' },
      ],
    },
    { name: 'columns', status: 'error', error: 'Unresolved placeholders, nothing was sent', operations: [], waits: [] },
    { name: 'names', status: 'skipped', operations: [], waits: [] },
  ],
};

test('text log lists stages, waits, placeholders and failures', () => {
  const text = formatRunText(run);
  assert.match(text, /Site:\s+https:\/\/contoso\.sharepoint\.com\/sites\/demo/);
  assert.match(text, /== Stage 1: lists \[completed\] 1 succeeded, 1 failed, 1 already existed/);
  assert.match(text, /Waited 7 s: 1 operation\(s\) throttled \(retry 1 of 3\)/);
  assert.match(text, /\[OK  \]   1 line 8    201 POST \/_api\/web\/lists \(attempt 2\)/);
  assert.match(text, /\[WARN\]   2 500 POST \/_api\/web\/lists\n\s+A list \.\.\. already exists/);
  assert.match(text, /\[FAIL\]   3 400 POST \/_api\/web\/lists\n\s+Boom/);
  assert.match(text, /== Stage 2: columns \[error\] not sent\nERROR: Unresolved placeholders/);
  assert.match(text, /== Stage 3: names \[skipped\] skipped/);
});

test('JSON export round-trips and holds no digest or cookie fields', () => {
  const json = formatRunJson(run);
  assert.deepEqual(JSON.parse(json), run);
  assert.doesNotMatch(json, /digest|cookie|FedAuth/i);
});

test('shortUrl strips the site prefix only', () => {
  assert.equal(shortUrl(SITE + '/_api/web', SITE), '/_api/web');
  assert.equal(shortUrl('https://other.sharepoint.com/_api/web', SITE), 'https://other.sharepoint.com/_api/web');
  assert.equal(shortUrl(null, SITE), '(no URL)');
});
