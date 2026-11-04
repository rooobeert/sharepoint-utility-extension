import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRunner, locateBatchError } from '../extension/lib/runner.js';
import { parseBatch } from '../extension/lib/batch.js';
import { readFixture, readSample } from './helpers.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const LIST_GUID = '0b1c2d3e-4f50-4a6b-8c7d-9e0f1a2b3c4d';
const WEB_GUID = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

const part = (status, text, body = '', extra = '') =>
  `--resp\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\nHTTP/1.1 ${status} ${text}\r\nCONTENT-TYPE: application/json;odata=nometadata\r\n${extra}\r\n${body}\r\n`;
const multipart = (...parts) => ({
  status: 200,
  statusText: 'OK',
  phase: 'request',
  headers: { 'content-type': 'multipart/mixed; boundary=resp' },
  text: parts.join('') + '--resp--\r\n',
});

const json = (status, body) => ({ status, statusText: '', headers: {}, text: JSON.stringify(body) });
const notFound = () => json(404, { 'odata.error': { message: { value: 'File Not Found.' } } });

// Fake page: GETs answer from `sp` (lists of the web, web, site collection),
// $batch calls answer from a queue.
function fakePage(batchReplies, sp = {}) {
  const sent = [];
  const gets = [];
  const request = async (req) => {
    if (req.method === 'GET') {
      gets.push(req.url);
      const path = req.url.replace(/^https:\/\/[^/]+(\/sites\/[^/]+)?/, '');
      if (path.startsWith('/_api/web/lists?')) return sp.lists ? json(200, { value: sp.lists }) : notFound();
      if (path.startsWith('/_api/web?')) return sp.web ? json(200, sp.web) : notFound();
      if (path.startsWith('/_api/site?')) return sp.site ? json(200, sp.site) : notFound();
      return notFound();
    }
    sent.push(req);
    return batchReplies.shift();
  };
  return { request, sent, gets };
}

function harness(replies, lookups) {
  const page = fakePage(replies, lookups);
  const sleeps = [];
  const events = [];
  const runner = createRunner({
    request: page.request,
    sleep: async (ms) => sleeps.push(ms),
    onEvent: (e) => events.push(e),
  });
  return { runner, page, sleeps, events };
}

const fixtureStage = { name: 'fixture', body: readFixture('request-mixed.txt') };
const fixtureLookups = {
  lists: [{ Id: LIST_GUID, Title: 'Sites', RootFolder: { ServerRelativeUrl: '/sites/demo/Lists/crSites' } }],
  web: { Id: WEB_GUID, Title: 'Demo', Url: SITE },
  site: { Url: SITE },
};

test('runStage: headers, CRLF body, digest requested, placeholders replaced', async () => {
  const { runner, page } = harness(
    [multipart(part(201, 'Created'), part(201, 'Created'), part(201, 'Created'), part(201, 'Created'), part(204, 'No Content'))],
    fixtureLookups
  );
  const result = await runner.runStage(SITE + '/', fixtureStage, 0);
  assert.equal(page.sent.length, 1);
  const req = page.sent[0];
  assert.equal(req.url, SITE + '/_api/$batch');
  assert.equal(req.digestFrom, SITE);
  assert.equal(req.headers['Content-Type'], 'multipart/mixed; boundary=batch_fixture');
  assert.equal(req.headers.Accept, 'application/json;odata=nometadata');
  assert.ok(!/[^\r]\n/.test(req.body));
  assert.ok(req.body.endsWith('--batch_fixture--\r\n'));
  assert.ok(!req.body.includes('{{') && !req.body.includes('objListGuids'));
  assert.ok(req.body.includes(`{${LIST_GUID}}`));
  assert.deepEqual(result.summary, { total: 5, succeeded: 5, warnings: 0, failed: 0, throttled: 0 });
});

test('runStage: unresolved placeholder stops before sending', async () => {
  const { runner, page } = harness([], { lists: [], web: { Id: WEB_GUID } });
  await assert.rejects(runner.runStage(SITE, fixtureStage, 0), /Unresolved placeholders.*crSites.*List not found/);
  assert.equal(page.sent.length, 0);
});

test('runStage: inner 429 re-sends only the throttled operation after Retry-After', async () => {
  const fixture = readFixture('response-mixed.txt').replace(/\n/g, '\r\n');
  const first = {
    status: 200,
    statusText: 'OK',
    headers: { 'content-type': 'multipart/mixed; boundary=batchresponse_5e1c9a2f-7b3d-4c11-9a0e-2f6d8b1c4a77' },
    text: fixture,
  };
  const { runner, page, sleeps } = harness([first, multipart(part(201, 'Created'))], fixtureLookups);
  const result = await runner.runStage(SITE, fixtureStage, 0);

  assert.deepEqual(sleeps, [7000]);
  assert.equal(page.sent.length, 2);
  const retried = parseBatch(page.sent[1].body);
  assert.equal(retried.operations.length, 1);
  assert.match(retried.operations[0].body, /crSite/);
  assert.equal(page.sent[1].digestFrom, SITE); // fresh digest for the retry
  assert.deepEqual(result.operations.map((o) => o.outcome), ['success', 'warning', 'success', 'failed', 'success']);
  assert.equal(result.operations[2].attempts, 2);
  assert.deepEqual(result.summary, { total: 5, succeeded: 3, warnings: 1, failed: 1, throttled: 0 });
  assert.equal(result.waits.length, 1);
});

test('runStage: batch-level 503 waits (default 10 s) and gives up after 3 retries', async () => {
  const busy = { status: 503, statusText: 'Service Unavailable', headers: {}, text: '' };
  const { runner, page, sleeps, events } = harness([busy, busy, busy, busy]);
  const stage = { name: 'lists', body: readSample('Batch1.txt') };
  await assert.rejects(runner.runStage('https://contoso.sharepoint.com/sites/rd-hub-projects', stage, 0), /HTTP 503/);
  assert.equal(page.sent.length, 4);
  assert.deepEqual(sleeps, [10000, 10000, 10000]);
  assert.equal(events.filter((e) => e.type === 'wait').length, 3);
});

test('runStage: digest failure is reported clearly', async () => {
  const { runner } = harness([
    { phase: 'contextinfo', status: 403, statusText: 'Forbidden', headers: {}, text: '{"odata.error":{"message":{"value":"Access denied."}}}' },
  ]);
  await assert.rejects(
    runner.runStage(SITE, { name: 'x', body: readSample('Batch1.txt') }, 0),
    /contextinfo.*403.*Access denied/
  );
});

test('runStages: pauses after a failing stage and skips the rest when told to stop', async () => {
  const ok = multipart(part(201, 'Created'), part(201, 'Created'));
  const bad = multipart(part(400, 'Bad Request', '{"odata.error":{"message":{"value":"Boom"}}}'), part(201, 'Created'));
  const stages = ['A', 'B', 'C'].map((name) => ({ name, body: readSample('ste-batch1-create-lists.txt') }));

  const asked = [];
  const stop = harness([ok, bad, ok]);
  const run = await stop.runner.runStages(SITE, stages, { shouldContinue: async (r, i) => (asked.push(i), false) });
  assert.deepEqual(asked, [1]);
  assert.deepEqual(run.stages.map((s) => s.status), ['completed', 'completed', 'skipped']);
  assert.equal(run.stages[1].summary.failed, 1);

  const go = harness([ok, bad, ok]);
  const run2 = await go.runner.runStages(SITE, stages, { shouldContinue: async () => true });
  assert.deepEqual(run2.stages.map((s) => s.status), ['completed', 'completed', 'completed']);
});

test('runStages: already-exists warnings do not pause the run', async () => {
  const dup = multipart(
    part(500, 'Internal Server Error', '{"odata.error":{"message":{"value":"A list, survey, discussion board, or document library with the specified title already exists in this Web site."}}}'),
    part(201, 'Created')
  );
  const stages = ['A', 'B'].map((name) => ({ name, body: readSample('ste-batch1-create-lists.txt') }));
  let asked = false;
  const h = harness([dup, dup]);
  const run = await h.runner.runStages(SITE, stages, { shouldContinue: async () => (asked = true) });
  assert.equal(asked, false);
  assert.deepEqual(run.stages.map((s) => s.summary.warnings), [1, 1]);
});

test('runStage: lists are read once per stage and the flow object is kept in the result', async () => {
  const ok = multipart(part(201, 'Created'), part(201, 'Created'), part(201, 'Created'), part(201, 'Created'), part(204, 'No Content'));
  const { runner, page } = harness([ok], fixtureLookups);
  const result = await runner.runStage(SITE, fixtureStage, 0);
  assert.equal(page.gets.filter((u) => u.includes('/_api/web/lists?')).length, 1);
  assert.deepEqual(result.listGuids, { '/sites/demo/Lists/crSites': LIST_GUID });
});

test('runStage: a structurally broken body is never sent', async () => {
  const { runner, page } = harness([]);
  const broken = { name: 'x', body: readSample('Batch1.txt').replace('--changeset_provisioning--\n', '') };
  await assert.rejects(runner.runStage(SITE, broken, 0), /invalid, nothing was sent: Changeset "changeset_provisioning" \(started at line 1\) is never closed/);
  assert.equal(page.sent.length, 0);
});

test('preview: no batch is sent, placeholders reported', async () => {
  const { runner, page } = harness([], { lists: [], web: { Id: WEB_GUID } });
  const p = await runner.preview(SITE, fixtureStage);
  assert.equal(page.sent.length, 0);
  assert.equal(p.operations.length, 5);
  assert.deepEqual(p.placeholders.map((x) => [x.label, x.id, x.error]), [
    ['List "/sites/demo/Lists/crSites"', null, 'List not found.'],
    ['{{WebId}}', WEB_GUID, null],
  ]);
});

test('verifySite: returns title, web URL and site collection URL from SharePoint', async () => {
  const { runner, page } = harness([], { web: { Title: 'Demo', Url: SITE + '/' }, site: { Url: SITE } });
  assert.deepEqual(await runner.verifySite(SITE + '/'), { title: 'Demo', webUrl: SITE, siteCollectionUrl: SITE });
  assert.deepEqual(page.gets, [SITE + '/_api/web?$select=Title,Url', SITE + '/_api/site?$select=Url']);
  assert.equal(page.sent.length, 0);
});

test('verifySite: a site that does not exist is an error', async () => {
  const { runner } = harness([], {});
  await assert.rejects(runner.verifySite(SITE), /Could not read the site .*HTTP 404: File Not Found/);
});

test('batch-level "invalid HTTP method" error names the operations and their lines', async () => {
  const msg = "An invalid HTTP method 'MERGE' was detected for a query operation. Query operations only support the HTTP 'GET' method.";
  const rejected = { status: 400, statusText: 'Bad Request', headers: {}, text: JSON.stringify({ 'odata.error': { message: { value: msg } } }) };
  const rd = '/sites/rd-hub-projects/Lists/';
  const lists = ['rdProjects', 'crProductionSites', 'crRequests', 'crDepartments'].map((n, i) => ({
    Id: `0000000${i}-0000-4000-8000-000000000000`,
    RootFolder: { ServerRelativeUrl: rd + n },
  }));
  const { runner } = harness([rejected], { lists });
  const err = await runner.runStage('https://contoso.sharepoint.com/sites/rd-hub-projects', { name: 'b2', body: readSample('Batch2.txt') }, 1).catch((e) => e);
  assert.match(err.message, /HTTP 400 Bad Request: An invalid HTTP method 'MERGE'/);
  assert.match(err.message, /MERGE operations in this stage: operation 1 at line 8\./);
  assert.match(err.message, /Check the lines before line 8/);
});

test('locateBatchError: generic batch error gives the line range', () => {
  const ops = parseBatch(readSample('Batch1.txt')).operations;
  assert.deepEqual(locateBatchError('Something odd.', ops), [
    'SharePoint rejected the whole batch, so no operation ran. The stage has 7 operation(s), lines 8–68.',
  ]);
});
