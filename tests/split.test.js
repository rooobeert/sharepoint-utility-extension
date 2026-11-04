import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRunner } from '../extension/lib/runner.js';
import { parseBatch } from '../extension/lib/batch.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const json = (status, body) => ({ status, statusText: '', headers: {}, text: JSON.stringify(body) });

/** One batch with a changeset; each op is [requestLine, body]. */
function batch(ops) {
  const lines = ['--batch_t', 'Content-Type: multipart/mixed; boundary=cs_t', ''];
  for (const [req, body, extra = []] of ops) {
    lines.push('--cs_t', 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '', req,
      'Content-Type: application/json;odata=verbose', 'Accept: application/json;odata=nometadata', ...extra, '', body, '');
  }
  lines.push('--cs_t--', '--batch_t--');
  return lines.join('\n');
}
const createList = (title, template = 100) => [
  `POST ${SITE}/_api/web/lists HTTP/1.1`,
  JSON.stringify({ __metadata: { type: 'SP.List' }, BaseTemplate: template, Title: title }),
];
const lookupField = (onList, name, ref) => [
  `POST ${SITE}/_api/web/getList('/sites/demo/Lists/${onList}')/fields/CreateFieldAsXml HTTP/1.1`,
  `{"parameters": {"SchemaXml": "<Field Type=\\"Lookup\\" Name=\\"${name}\\" List=\\"{${ref}}\\" ShowField=\\"Title\\" />", "Options": 24}}`,
];
const textField = (onList, name) => [
  `POST ${SITE}/_api/web/getList('/sites/demo/Lists/${onList}')/fields/CreateFieldAsXml HTTP/1.1`,
  `{"parameters": {"SchemaXml": "<Field Type=\\"Text\\" Name=\\"${name}\\" />", "Options": 24}}`,
];

/**
 * Stateful fake SharePoint: creating a list in a batch adds it, so a later
 * lookup finds it. `fail` maps a list title to an error message for its creation.
 */
function fakeSharePoint({ lists = [], fail = {} } = {}) {
  let n = 0;
  const state = { lists: [...lists], batches: [], gets: [] };
  const request = async (req) => {
    if (req.method === 'GET') {
      state.gets.push(req.url);
      const path = req.url.slice(SITE.length);
      if (path.startsWith('/_api/web/lists?')) {
        return json(200, { value: state.lists.map((l) => ({ Id: l.id, Title: l.title, RootFolder: { ServerRelativeUrl: l.url } })) });
      }
      if (path.startsWith('/_api/web?')) return json(200, { Id: 'web-guid', ServerRelativeUrl: '/sites/demo', Url: SITE, Title: 'Demo' });
      if (path.startsWith('/_api/site?')) return json(200, { Url: SITE });
      return json(404, {});
    }
    state.batches.push(req.body);
    const ops = parseBatch(req.body).operations;
    let out = '';
    for (const op of ops) {
      let res = 'HTTP/1.1 201 Created\r\n\r\n{}';
      if (/\/_api\/web\/lists$/.test(op.url)) {
        const { Title, BaseTemplate } = JSON.parse(op.body);
        if (fail[Title]) {
          res = `HTTP/1.1 400 Bad Request\r\n\r\n${JSON.stringify({ 'odata.error': { message: { value: fail[Title] } } })}`;
        } else {
          n++;
          state.lists.push({ id: `0000000${n}-0000-4000-8000-000000000000`, title: Title, url: BaseTemplate === 101 ? `/sites/demo/${Title}` : `/sites/demo/Lists/${Title}` });
        }
      }
      out += `--r\r\nContent-Type: application/http\r\n\r\n${res}\r\n`;
    }
    return { status: 200, statusText: 'OK', headers: { 'content-type': 'multipart/mixed; boundary=r' }, text: out + '--r--\r\n' };
  };
  return { request, state };
}

const runnerFor = (sp) => {
  const events = [];
  return { events, runner: createRunner({ request: sp.request, sleep: async () => {}, onEvent: (e) => events.push(e) }) };
};

test('a list created in the same stage: sent in 2 parts, lookup filled in between, order kept', async () => {
  const body = batch([
    createList('crSites'),
    textField('crRequests', 'crTitle2'),
    lookupField('crRequests', 'crSite', '{{ListId:Lists/crSites}}'),
    textField('crRequests', 'crAfter'),
  ]);
  const sp = fakeSharePoint({ lists: [{ id: 'req-guid', title: 'crRequests', url: '/sites/demo/Lists/crRequests' }] });
  const { runner, events } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 'one file', body }, 0);

  assert.equal(sp.state.batches.length, 2);
  assert.deepEqual(parseBatch(sp.state.batches[0]).operations.map((o) => o.method + ' ' + o.url.split('/').pop()), ['POST lists', 'POST CreateFieldAsXml']);
  const second = parseBatch(sp.state.batches[1]).operations;
  assert.equal(second.length, 2);
  assert.match(second[0].body, /List=\\"\{00000001-0000-4000-8000-000000000000\}\\"/);
  assert.equal(result.parts, 2);
  assert.deepEqual(result.operations.map((o) => [o.index, o.outcome]), [[0, 'success'], [1, 'success'], [2, 'success'], [3, 'success']]);
  assert.deepEqual(result.operations.map((o) => o.line), parseBatch(body).operations.map((o) => o.line));
  assert.match(result.notes[0], /Sent in 2 parts: .* before line 28\./);
  assert.ok(events.some((e) => e.type === 'part' && e.part === 2 && e.from === 28));
});

test('one stage, two dependency levels: three parts', async () => {
  const body = batch([
    createList('A'),
    createList('B'),
    lookupField('B', 'toA', '{{ListId:A}}'),
    createList('C'),
    lookupField('C', 'toC', '{{ListId:Lists/C}}'),
  ]);
  const sp = fakeSharePoint();
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  // A and B in part 1; the lookup to A and creating C in part 2; the lookup to C in part 3.
  assert.deepEqual(sp.state.batches.map((b) => parseBatch(b).operations.length), [2, 2, 1]);
  assert.equal(result.parts, 3);
  assert.equal(result.summary.succeeded, 5);
});

test('lists that already exist: one batch, sent byte-for-byte with IDs filled in', async () => {
  const body = batch([lookupField('crRequests', 'crSite', "@{variables('anyName')?['/sites/demo/Lists/crSites']}")]);
  const sp = fakeSharePoint({ lists: [{ id: 'site-guid', title: 'Sites', url: '/sites/demo/Lists/crSites' }] });
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  assert.equal(sp.state.batches.length, 1);
  assert.equal(sp.state.batches[0], parseBatch(body).body.replace("@{variables('anyName')?['/sites/demo/Lists/crSites']}", 'site-guid'));
  assert.equal(result.parts, 1);
  assert.deepEqual(result.notes, []);
});

test('the list creation fails: stops after that part, explains, rest not sent', async () => {
  const body = batch([createList('crSites'), lookupField('crRequests', 'crSite', '{{ListId:crSites}}'), textField('crRequests', 'x')]);
  const sp = fakeSharePoint({ fail: { crSites: 'Access denied.' } });
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  assert.equal(sp.state.batches.length, 1);
  assert.match(result.error, /Operation 2 \(line 18\) needs List "crSites" \(List not found\.\) It should have been created by operation 1 \(line 8\), which failed: Access denied\./);
  assert.deepEqual(result.operations.map((o) => o.outcome), ['failed', 'failed', 'failed']);
  assert.match(result.operations[2].errorMessage, /^Not sent: a lookup needed at line 18/);
  assert.equal(result.summary.failed, 3);
});

test('a list that nothing creates: stops before sending anything', async () => {
  const body = batch([textField('crRequests', 'x'), lookupField('crRequests', 'crSite', '{{ListId:Lists/missing}}')]);
  const sp = fakeSharePoint();
  const { runner } = runnerFor(sp);
  await assert.rejects(
    runner.runStage(SITE, { name: 's', body }, 0),
    /nothing was sent: Operation 2 \(line 18\) needs List "Lists\/missing" \(List not found\.\) No earlier operation in this stage creates it\./
  );
  assert.equal(sp.state.batches.length, 0);
});

test('a library created in the same stage, referenced by its URL name', async () => {
  const body = batch([createList('crDocumentVault', 101), lookupField('crRequests', 'doc', '{{ListId:crDocumentVault}}')]);
  const sp = fakeSharePoint();
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  assert.equal(result.parts, 2);
  assert.equal(result.summary.succeeded, 2);
});

test('URL placeholders: filled in from the verified site, line numbers unchanged', async () => {
  const body = batch([
    [`POST {{WebUrl}}/_api/web/getList('{{WebServerRelativeUrl}}/Lists/crRequests')/fields/CreateFieldAsXml HTTP/1.1`, '{"parameters": {"SchemaXml": "<Field Type=\\"Text\\" Name=\\"x\\" />"}}'],
  ]);
  const sp = fakeSharePoint();
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  const sent = parseBatch(sp.state.batches[0]).operations[0];
  assert.equal(sent.url, `${SITE}/_api/web/getList('/sites/demo/Lists/crRequests')/fields/CreateFieldAsXml`);
  assert.equal(result.operations[0].line, 8);
  assert.deepEqual(result.placeholders.map((p) => [p.label, p.id]), [
    ['{{WebUrl}}', SITE],
    ['{{WebServerRelativeUrl}}', '/sites/demo'],
  ]);
});

test('preview plans the parts and flags a list used before it is created', async () => {
  const sp = fakeSharePoint();
  const { runner } = runnerFor(sp);
  const ok = await runner.preview(SITE, { name: 's', body: batch([createList('A'), lookupField('X', 'toA', '{{ListId:A}}')]) });
  assert.equal(ok.parts, 2);
  assert.deepEqual(ok.placeholders.map((p) => [p.label, p.id, p.pending]), [['List "A"', null, { line: 8, operation: 1 }]]);
  assert.match(ok.notes[0], /Sent in 2 parts/);
  assert.equal(sp.state.batches.length, 0);

  const wrong = await runner.preview(SITE, { name: 's', body: batch([lookupField('X', 'toA', '{{ListId:A}}'), createList('A')]) });
  assert.ok(wrong.warnings.some((w) => /List "A" is used at line 8 before it is created at line 18/.test(w)));
});

test('SharePoint rejects part 2: part 1 results are kept', async () => {
  const body = batch([createList('A'), lookupField('X', 'toA', '{{ListId:A}}')]);
  const sp = fakeSharePoint();
  const inner = sp.request;
  let posts = 0;
  sp.request = async (req) => {
    if (req.method === 'POST' && ++posts === 2) return json(400, { 'odata.error': { message: { value: 'Bad batch.' } } });
    return inner(req);
  };
  const { runner } = runnerFor(sp);
  const result = await runner.runStage(SITE, { name: 's', body }, 0);
  assert.deepEqual(result.operations.map((o) => o.outcome), ['success', 'failed']);
  assert.match(result.error, /HTTP 400.*Bad batch\. .*Earlier parts were sent/);
});
