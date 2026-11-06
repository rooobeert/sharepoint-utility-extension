import { test } from 'node:test';
import assert from 'node:assert/strict';
import { jsonRule, pathsRegex, addPath, normalizePath, DEFAULT_PATHS } from '../extension/lib/json-header.js';

test('rule sets Accept for main-frame page loads only', () => {
  const r = jsonRule();
  assert.deepEqual(r.condition.resourceTypes, ['main_frame']);
  assert.equal(r.action.requestHeaders[0].header, 'Accept');
  assert.match(r.action.requestHeaders[0].value, /^application\/json/);
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

test('an empty list gives no rule', () => {
  assert.equal(jsonRule([]), null);
});
