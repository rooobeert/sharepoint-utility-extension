import { test } from 'node:test';
import assert from 'node:assert/strict';
import { headerRule, pathsRegex, addPath, normalizePath, normalizeHeader, setHeader, DEFAULT_PATHS, DEFAULT_HEADERS } from '../extension/lib/request-headers.js';

test('default rule sets Accept JSON for main-frame page loads only', () => {
  const r = headerRule();
  assert.deepEqual(r.condition.resourceTypes, ['main_frame']);
  assert.deepEqual(r.action.requestHeaders, [{ header: 'Accept', operation: 'set', value: 'application/json;odata=nometadata' }]);
});

test('every listed header is set', () => {
  const r = headerRule(['/_api/'], [...DEFAULT_HEADERS, { name: 'Accept-Language', value: 'de-DE' }]);
  assert.deepEqual(r.action.requestHeaders.map((h) => h.header), ['Accept', 'Accept-Language']);
  assert.equal(r.action.requestHeaders[1].value, 'de-DE');
});

test('regex matches SharePoint URLs whose path contains a listed text', () => {
  const re = new RegExp(pathsRegex(DEFAULT_PATHS), 'i');
  assert.ok(re.test('https://t.sharepoint.com/sites/x/_api/web/lists?$select=Title'));
  assert.ok(re.test('https://t.sharepoint.com/_api/web'));
  assert.ok(!re.test('https://t.sharepoint.com/sites/x/SitePages/_api.aspx'));
  assert.ok(!re.test('https://t.sharepoint.com/sites/x/page.aspx?next=/_api/web'));
  assert.ok(!re.test('https://evil.example.com/sites/x/_api/web'));
  assert.ok(!re.test('https://t.sharepoint.com.evil.com/_api/web'));
  assert.ok(!re.test('http://t.sharepoint.com/_api/web'));
});

test('several paths are alternatives; special characters are escaped', () => {
  const re = new RegExp(pathsRegex(['/_api/', '/_vti_bin/client.svc']), 'i');
  assert.ok(re.test('https://t.sharepoint.com/sites/x/_vti_bin/client.svc/web'));
  assert.ok(!re.test('https://t.sharepoint.com/sites/x/_vti_bin/clientXsvc/web'));
  assert.ok(re.test('https://t.sharepoint.com/_API/web'));
});

test('addPath trims, rejects unusable and duplicate paths', () => {
  assert.deepEqual(addPath(['/_api/'], ' /_vti_bin/ '), ['/_api/', '/_vti_bin/']);
  assert.deepEqual(addPath(['/_api/'], '/_API/'), ['/_api/']);
  assert.equal(normalizePath('a b'), null);
  assert.equal(normalizePath('?x'), null);
  assert.equal(normalizePath('  '), null);
});

test('headers: validation and replace by name', () => {
  assert.deepEqual(normalizeHeader(' Accept-Language ', ' de-DE '), { name: 'Accept-Language', value: 'de-DE' });
  assert.equal(normalizeHeader('Bad Name', 'x'), null);
  assert.equal(normalizeHeader('X-Test', ''), null);
  assert.equal(normalizeHeader('X-Test', 'a\r\nb: c'), null);
  const list = setHeader(DEFAULT_HEADERS, 'accept', 'application/xml');
  assert.deepEqual(list, [{ name: 'accept', value: 'application/xml' }]);
  assert.equal(setHeader(list, 'X-A', 'v').length, 2);
  assert.equal(setHeader(list, 'bad name', 'v'), list);
});

test('an empty list gives no rule', () => {
  assert.equal(headerRule([], DEFAULT_HEADERS), null);
  assert.equal(headerRule(['/_api/'], []), null);
});
