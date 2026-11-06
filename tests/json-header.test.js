import { test } from 'node:test';
import assert from 'node:assert/strict';
import { jsonRule, JSON_URL_REGEX } from '../extension/lib/json-header.js';

test('rule sets Accept for main-frame page loads only', () => {
  const r = jsonRule();
  assert.deepEqual(r.condition.resourceTypes, ['main_frame']);
  assert.equal(r.action.requestHeaders[0].header, 'Accept');
  assert.match(r.action.requestHeaders[0].value, /^application\/json/);
});

test('regex matches SharePoint _api URLs on sharepoint.com only', () => {
  const re = new RegExp(JSON_URL_REGEX);
  assert.ok(re.test('https://t.sharepoint.com/sites/x/_api/web/lists?$select=Title'));
  assert.ok(re.test('https://t.sharepoint.com/_api/web'));
  assert.ok(!re.test('https://t.sharepoint.com/sites/x/SitePages/_api.aspx'));
  assert.ok(!re.test('https://evil.example.com/sites/x/_api/web'));
  assert.ok(!re.test('https://t.sharepoint.com.evil.com/_api/web'));
  assert.ok(!re.test('http://t.sharepoint.com/_api/web'));
});
