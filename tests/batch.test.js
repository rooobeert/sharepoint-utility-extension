import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  CRLF,
  normalizeLineEndings,
  detectBoundary,
  prepareBatchBody,
  parseBatch,
  buildBatch,
  analyseBatch,
  siteOfOperationUrl,
  targetSiteFromStages,
  siteUrlFromPageUrl,
  isSharePointUrl,
  isOnOrigin,
} from '../extension/lib/batch.js';
import { readSample, readFixture, sampleNames } from './helpers.js';

// Operation counts per sample = number of "HTTP/1.1" request lines in the file.
const countRequestLines = (text) => (text.match(/ HTTP\/1\.1\r?$/gm) || []).length;

test('every sample parses into the expected number of operations', () => {
  const names = sampleNames();
  assert.ok(names.length > 0, 'samples folder is empty');
  for (const name of names) {
    const text = readSample(name);
    const parsed = parseBatch(text);
    assert.equal(parsed.operations.length, countRequestLines(text), name);
    assert.deepEqual(parsed.errors, [], `${name}: ${parsed.errors.join(' | ')}`);
    assert.deepEqual(parsed.warnings, [], `${name}: ${parsed.warnings.join(' | ')}`);
    for (const op of parsed.operations) {
      assert.match(op.method, /^(POST|MERGE|PATCH|GET|PUT|DELETE)$/, name);
      assert.match(op.url, /^https:\/\/[^/]+\.sharepoint\.com\/.+\/_api\//, name);
      assert.ok(op.changeset, `${name}: operation outside a changeset`);
    }
  }
});

test('boundaries are detected from the first line', () => {
  assert.equal(detectBoundary(readSample('Batch1.txt')), 'batch_provisioning');
  assert.equal(detectBoundary(readSample('ste-batch1-create-lists.txt')), 'batch_hrp');
  assert.equal(detectBoundary('\n\n--batch_x\nfoo'), 'batch_x');
  assert.throws(() => detectBoundary('POST https://x HTTP/1.1'), /boundary/);
});

test('Batch1: list creation operations are read correctly', () => {
  const { operations, boundary } = parseBatch(readSample('Batch1.txt'));
  assert.equal(boundary, 'batch_provisioning');
  assert.equal(operations.length, 7);
  const first = operations[0];
  assert.equal(first.method, 'POST');
  assert.equal(first.url, 'https://contoso.sharepoint.com/sites/rd-hub-projects/_api/web/lists');
  assert.equal(first.changeset, 'changeset_provisioning');
  assert.equal(first.headers['content-type'], 'application/json;odata=verbose');
  assert.equal(JSON.parse(first.body).Title, 'crRequests');
  assert.equal(JSON.parse(operations[3].body).BaseTemplate, 101);
});

test('Batch4: MERGE headers (IF-MATCH, Accept-Language) are kept', () => {
  const { operations } = parseBatch(readSample('Batch4.txt'));
  assert.equal(operations.length, 14);
  assert.equal(operations[0].method, 'MERGE');
  assert.equal(operations[0].headers['if-match'], '*');
  assert.equal(operations[0].headers['accept-language'], 'de-DE');
  assert.equal(operations[13].headers['accept-language'], 'en-US');
});

test('prepareBatchBody: LF becomes CRLF, closing boundary and trailing CRLF', () => {
  for (const name of sampleNames()) {
    const { body, boundary } = prepareBatchBody(readSample(name));
    assert.ok(!/[^\r]\n/.test(body), `${name}: bare LF left`);
    assert.ok(body.endsWith(`--${boundary}--${CRLF}`), `${name}: not closed`);
  }
  // ste-* files have no trailing newline after the closing boundary
  const ste = prepareBatchBody(readSample('ste-batch1-create-lists.txt')).body;
  assert.ok(ste.endsWith('--cs_d4b1379312c8--\r\n--batch_hrp--\r\n'));
});

test('missing closing batch boundary is an error and is not patched', () => {
  const text = readSample('Batch1.txt').replace(/--batch_provisioning--\s*$/, '');
  assert.ok(prepareBatchBody(text).body.endsWith('--changeset_provisioning--\r\n'));
  const parsed = parseBatch(text);
  assert.equal(parsed.operations.length, 7);
  assert.deepEqual(parsed.errors, ['The batch is never closed: the closing line "--batch_provisioning--" is missing (the text ends at line 75).']);
});

test('prepareBatchBody does not change content other than line endings', () => {
  const text = readSample('Batch2.txt');
  const { body } = prepareBatchBody(text);
  assert.equal(body.replace(/\r\n/g, '\n').trimEnd(), text.trimEnd());
});

test('mixed CR / LF / CRLF input is normalised', () => {
  assert.equal(normalizeLineEndings('﻿a\nb\r\nc\rd'), 'a\r\nb\r\nc\r\nd');
});

test('unclosed changeset is an error', () => {
  const text = readSample('Batch1.txt').replace('--changeset_provisioning--\n', '');
  const parsed = parseBatch(text);
  assert.equal(parsed.operations.length, 7);
  assert.deepEqual(parsed.errors, ['Changeset "changeset_provisioning" (started at line 1) is never closed: the closing line "--changeset_provisioning--" is missing.']);
});

test('unreadable request line and empty batch are errors', () => {
  const bad = parseBatch(readSample('Batch1.txt').replace('POST https://contoso', 'POSTX'));
  assert.ok(bad.errors.some((e) => /^Operation 1 \(line 8\): could not read the request line/.test(e)), bad.errors.join());
  assert.deepEqual(parseBatch('--batch_a\n--batch_a--').errors, ['No operations found in this batch.']);
});

test('every operation knows the line of its request line, also with leading blank lines', () => {
  for (const name of sampleNames()) {
    const text = readSample(name);
    const lines = text.split('\n');
    for (const op of parseBatch(text).operations) {
      assert.ok(lines[op.line - 1].startsWith(`${op.method} `), `${name} op ${op.index + 1} -> line ${op.line}`);
    }
  }
  const shifted = parseBatch('\n\n' + readSample('Batch1.txt'));
  assert.deepEqual(shifted.operations.slice(0, 2).map((o) => o.line), [10, 20]);
  assert.ok(shifted.body.startsWith('--batch_provisioning')); // sent body is still trimmed
});

test('write operations outside a changeset are errors with their line', () => {
  // Remove the changeset wrapper: operations become direct batch parts.
  const flat = readSample('Batch4.txt')
    .replace('Content-Type: multipart/mixed; boundary=changeset_provisioning\n\n--changeset_provisioning\n', '')
    .replace(/--changeset_provisioning--\n/, '')
    .replace(/^--changeset_provisioning$/gm, '--batch_provisioning');
  const parsed = parseBatch(flat);
  assert.equal(parsed.operations.length, 14);
  assert.equal(parsed.errors.length, 14);
  assert.match(parsed.errors[0], /^Operation 1 \(line 5\): MERGE is outside a changeset/);
  assert.equal(flat.split('\n')[4].slice(0, 6), 'MERGE ');
});

test('operations outside a changeset are supported', () => {
  const text = [
    '--batch_a',
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    '',
    'GET https://contoso.sharepoint.com/sites/demo/_api/web/lists HTTP/1.1',
    'Accept: application/json;odata=nometadata',
    '',
    '--batch_a--',
  ].join('\n');
  const { operations } = parseBatch(text);
  assert.equal(operations.length, 1);
  assert.equal(operations[0].method, 'GET');
  assert.equal(operations[0].changeset, null);
});

test('buildBatch round-trips all operations and can build a subset', () => {
  for (const name of ['Batch1.txt', 'ste-batch2-create-columns.txt', 'Batch4.txt']) {
    const parsed = parseBatch(readSample(name));
    const rebuilt = buildBatch(parsed.boundary, parsed.operations);
    const again = parseBatch(rebuilt);
    assert.equal(again.operations.length, parsed.operations.length, name);
    assert.deepEqual(again.operations.map((o) => o.raw), parsed.operations.map((o) => o.raw), name);
  }
  const parsed = parseBatch(readFixture('request-mixed.txt'));
  const subset = buildBatch(parsed.boundary, [parsed.operations[2]]);
  assert.ok(subset.startsWith('--batch_fixture\r\nContent-Type: multipart/mixed; boundary=changeset_fixture\r\n\r\n--changeset_fixture\r\n'));
  assert.ok(subset.endsWith('--changeset_fixture--\r\n--batch_fixture--\r\n'));
  const reparsed = parseBatch(subset);
  assert.equal(reparsed.operations.length, 1);
  assert.match(reparsed.operations[0].body, /crSite/);
});

test('analyseBatch warns about the 100-operation limit only', () => {
  assert.deepEqual(analyseBatch(parseBatch(readSample('Batch2.txt'))), []);
  assert.ok(analyseBatch(parseBatch(readSample('Batch3.txt'))).some((w) => /132 operations/.test(w)));
  assert.ok(analyseBatch(parseBatch(readSample('Batch5-other-tenant.txt'))).some((w) => /108 operations/.test(w)));
});

test('siteOfOperationUrl takes everything before the first /_api/', () => {
  assert.equal(
    siteOfOperationUrl("https://contoso.sharepoint.com/sites/rd-hub-projects/_api/web/getList('/sites/x')/fields/CreateFieldAsXml"),
    'https://contoso.sharepoint.com/sites/rd-hub-projects'
  );
  assert.equal(siteOfOperationUrl('https://contoso.sharepoint.com/_api/web/lists'), 'https://contoso.sharepoint.com');
  assert.equal(siteOfOperationUrl('https://contoso.sharepoint.com/sites/a/sub/_api/web'), 'https://contoso.sharepoint.com/sites/a/sub');
  assert.equal(siteOfOperationUrl('/_api/web/lists'), null);
  assert.equal(siteOfOperationUrl(null), null);
});

const stagesOf = (...names) => names.map((name) => ({ name, parsed: parseBatch(readSample(name)) }));

test('targetSiteFromStages: one site across all rd-hub stages', () => {
  const t = targetSiteFromStages(stagesOf('Batch1.txt', 'Batch2.txt', 'Batch3.txt', 'Batch4.txt'));
  assert.equal(t.site, 'https://contoso.sharepoint.com/sites/rd-hub-projects');
  assert.equal(t.relativeOnly, false);
  assert.deepEqual(t.sites[0].uses.map((u) => [u.stage, u.operations.length]), [[0, 7], [1, 66], [2, 132], [3, 14]]);
});

test('targetSiteFromStages: different sites are a conflict with details', () => {
  const t = targetSiteFromStages(stagesOf('Batch1.txt', 'Batch5-other-tenant.txt', 'ste-batch1-create-lists.txt'));
  assert.equal(t.site, null);
  assert.equal(t.relativeOnly, false);
  assert.deepEqual(t.sites.map((s) => s.site), [
    'https://contoso.sharepoint.com/sites/rd-hub-projects',
    'https://fabrikam.sharepoint.com/sites/ops-use-cases', // different tenant
    'https://contoso.sharepoint.com/sites/hr-processes-app-dev',
  ]);
  assert.deepEqual(t.sites[1].uses, [{ stage: 1, name: 'Batch5-other-tenant.txt', operations: t.sites[1].uses[0].operations }]);
  assert.equal(t.sites[1].uses[0].operations.length, 108);
});

test('targetSiteFromStages: site URL casing differences are the same site', () => {
  const lower = parseBatch(readSample('Batch1.txt').replace(/rd-hub-projects\/_api/g, 'RD-HUB-PROJECTS/_api'));
  const t = targetSiteFromStages([{ name: 'a', parsed: lower }, ...stagesOf('Batch4.txt')]);
  assert.ok(t.site);
});

test('targetSiteFromStages: relative URLs only, or no stages', () => {
  const rel = parseBatch(readSample('Batch1.txt').replace(/https:\/\/contoso\.sharepoint\.com\/sites\/rd-hub-projects/g, ''));
  assert.deepEqual(targetSiteFromStages([{ name: 'r', parsed: rel }]), { site: null, sites: [], relativeOnly: true });
  assert.deepEqual(targetSiteFromStages([]), { site: null, sites: [], relativeOnly: true });
});

test('site URL fallback parsing from a page URL', () => {
  assert.equal(
    siteUrlFromPageUrl('https://contoso.sharepoint.com/sites/rd-hub-projects/Lists/crRequests/AllItems.aspx?x=1'),
    'https://contoso.sharepoint.com/sites/rd-hub-projects'
  );
  assert.equal(siteUrlFromPageUrl('https://contoso.sharepoint.com/teams/hr/SitePages/Home.aspx'), 'https://contoso.sharepoint.com/teams/hr');
  assert.equal(siteUrlFromPageUrl('https://contoso.sharepoint.com/SitePages/Home.aspx'), 'https://contoso.sharepoint.com');
  assert.equal(siteUrlFromPageUrl('not a url'), null);
});

test('only https *.sharepoint.com URLs are accepted', () => {
  assert.ok(isSharePointUrl('https://contoso.sharepoint.com/sites/a'));
  assert.ok(isSharePointUrl('https://contoso-my.sharepoint.com/personal/x'));
  assert.ok(!isSharePointUrl('http://contoso.sharepoint.com/sites/a'));
  assert.ok(!isSharePointUrl('https://contoso.sharepoint.com.evil.example/sites/a'));
  assert.ok(!isSharePointUrl('https://evil.example/?sharepoint.com'));
});

test('isOnOrigin: only absolute URLs on exactly the same origin', () => {
  const o = 'https://contoso.sharepoint.com';
  assert.ok(isOnOrigin('https://contoso.sharepoint.com/sites/a/_api/web', o));
  assert.ok(!isOnOrigin('https://other.sharepoint.com/_api/web', o));
  assert.ok(!isOnOrigin('https://contoso.sharepoint.com.evil.example/x', o));
  assert.ok(!isOnOrigin('http://contoso.sharepoint.com/x', o));
  assert.ok(!isOnOrigin('https://contoso.sharepoint.com:8443/x', o));
  assert.ok(!isOnOrigin('/_api/web', o));
});
