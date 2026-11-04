import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  findPlaceholders,
  resolvePlaceholders,
  applyPlaceholders,
  listRefKind,
  buildListGuids,
  findCreator,
  findUrlPlaceholders,
  applyUrlPlaceholders,
  urlPlaceholderValues,
  LISTS_PATH,
  WEB_PATH,
} from '../extension/lib/placeholders.js';
import { parseBatch } from '../extension/lib/batch.js';
import { readSample, readFixture, sampleNames } from './helpers.js';

const GUID_SITES = '0b1c2d3e-4f50-4a6b-8c7d-9e0f1a2b3c4d';
const GUID_WEB = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';

// What GET /_api/web/lists?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder returns.
const LISTS = {
  value: [
    { Id: GUID_SITES, Title: 'Production Sites', RootFolder: { ServerRelativeUrl: '/sites/demo/Lists/crSites' } },
    { Id: '{99999999-8888-4777-8666-555555555555}', Title: 'Documents', RootFolder: { ServerRelativeUrl: '/sites/demo/Shared Documents' } },
  ],
};

// Fake SharePoint: answers by path and records the calls.
function fakeFetch(table) {
  const calls = [];
  const fn = async (path) => {
    calls.push(path);
    return table[path] || { status: 404, json: null };
  };
  fn.calls = calls;
  return fn;
}

const sharePoint = () =>
  fakeFetch({
    [LISTS_PATH]: { status: 200, json: LISTS },
    [WEB_PATH]: { status: 200, json: { Id: '{' + GUID_WEB + '}', ServerRelativeUrl: '/sites/demo' } },
  });

test('finds {{WebId}}, {{ListId:...}} and the Power Automate alias', () => {
  const text = readFixture('request-mixed.txt');
  const found = findPlaceholders(text);
  assert.deepEqual(found.map((p) => p.key), ['list:url:/sites/demo/lists/crsites', 'web']);
  // Both syntaxes point at the same list, so one lookup with two tokens.
  assert.deepEqual(found[0].tokens, [
    '{{ListId:/sites/demo/Lists/crSites}}',
    "@{variables('objListGuids')?['/sites/demo/Lists/crSites']}",
  ]);
});

test('titles, URLs and whitespace variants', () => {
  const found = findPlaceholders('{{ ListId : My List }} {{ListId:my list}} {{ListId:https://contoso.sharepoint.com/sites/a/Lists/X/}} {{WebId}} {{ WebId }}');
  assert.equal(found.length, 3);
  assert.equal(found[0].ref, 'My List');
  assert.deepEqual(found[0].tokens, ['{{ ListId : My List }}', '{{ListId:my list}}']);
  assert.equal(found[1].ref, '/sites/a/Lists/X');
  assert.deepEqual(found[2].tokens, ['{{WebId}}', '{{ WebId }}']);
  assert.equal(listRefKind('Documents'), 'name');
  assert.equal(listRefKind('Lists/crSites'), 'relative');
  assert.equal(listRefKind('/sites/a/Shared Documents'), 'url');
  assert.equal(listRefKind('https://contoso.sharepoint.com/sites/a'), 'url');
});

test('Power Automate form works with any variable name; case-insensitive names', () => {
  const found = findPlaceholders(
    "@{variables('objListGuids')?['/sites/demo/Lists/crSites']} @{variables('listIds')?['Lists/crSites']} {{listid:Production Sites}} {{webid}}"
  );
  assert.deepEqual(found.map((p) => [p.type, p.ref]), [
    ['list', '/sites/demo/Lists/crSites'],
    ['list', 'Lists/crSites'],
    ['list', 'Production Sites'],
    ['web', ''],
  ]);
});

test('buildListGuids builds the flow object: server-relative URL -> id', () => {
  const g = buildListGuids(LISTS);
  assert.deepEqual(g.byUrl, {
    '/sites/demo/Lists/crSites': GUID_SITES,
    '/sites/demo/Shared Documents': '99999999-8888-4777-8666-555555555555',
  });
  assert.equal(g.byTitle['production sites'], GUID_SITES);
  assert.deepEqual(buildListGuids({ d: { results: LISTS.value } }).byUrl, g.byUrl); // odata=verbose shape
  assert.deepEqual(buildListGuids(null), { byUrl: {}, byTitle: {} });
});

test('resolve and apply: lists read once, braces kept only where the body has them', async () => {
  const text = readFixture('request-mixed.txt');
  const found = findPlaceholders(text);
  const fetchJson = sharePoint();
  const { resolved, listGuids } = await resolvePlaceholders(found, fetchJson);
  assert.deepEqual(fetchJson.calls, [WEB_PATH, LISTS_PATH]);
  assert.equal(listGuids['/sites/demo/Lists/crSites'], GUID_SITES);
  assert.deepEqual(resolved.map((r) => [r.id, r.error]), [[GUID_SITES, null], [GUID_WEB, null]]);

  const out = applyPlaceholders(text, resolved);
  assert.deepEqual(findPlaceholders(out), []);
  assert.ok(out.includes(`List=\\"{${GUID_SITES}}\\" ShowField`)); // {{{ListId:..}}} -> {guid}
  assert.ok(out.includes(`List=\\"{${GUID_SITES}}\\" />`)); // {@{variables(..)}} -> {guid}
  assert.ok(out.includes(`Change Requests ${GUID_WEB}`)); // no braces
  assert.equal(parseBatch(out).operations.length, 5);
});

test('lists are matched by URL (any case, full URL, trailing slash, spaces) and by title', async () => {
  const found = findPlaceholders(
    '{{ListId:/SITES/demo/lists/crsites}} {{ListId:https://contoso.sharepoint.com/sites/demo/Shared%20Documents/}} {{ListId:production sites}} {{ListId:Documents}}'
  );
  const fetchJson = sharePoint();
  const { resolved } = await resolvePlaceholders(found, fetchJson);
  assert.deepEqual(fetchJson.calls, [WEB_PATH, LISTS_PATH]); // web read for the title/name references
  assert.deepEqual(resolved.map((r) => r.id), [
    GUID_SITES,
    '99999999-8888-4777-8666-555555555555',
    GUID_SITES,
    '99999999-8888-4777-8666-555555555555',
  ]);
});

test('site-relative references: Lists/x, a library name, ./prefix', async () => {
  const found = findPlaceholders('{{ListId:Lists/crSites}} {{ListId:./Lists/CRSITES/}} {{ListId:Shared Documents}} {{ListId:Lists/nope}}');
  const { resolved } = await resolvePlaceholders(found, sharePoint());
  assert.equal(found.length, 3); // the first two are the same list
  assert.deepEqual(resolved.map((r) => r.id || r.error), [GUID_SITES, '99999999-8888-4777-8666-555555555555', 'List not found.']);
});

test('only URL references: the web is not read', async () => {
  const fetchJson = sharePoint();
  await resolvePlaceholders(findPlaceholders('{{ListId:/sites/demo/Lists/crSites}}'), fetchJson);
  assert.deepEqual(fetchJson.calls, [LISTS_PATH]);
});

test('a name matching one list by title and another by URL is ambiguous', async () => {
  const lists = {
    value: [
      { Id: 'aaaaaaaa-0000-4000-8000-000000000001', Title: 'Archive', RootFolder: { ServerRelativeUrl: '/sites/demo/Lists/old' } },
      { Id: 'aaaaaaaa-0000-4000-8000-000000000002', Title: 'Other', RootFolder: { ServerRelativeUrl: '/sites/demo/Archive' } },
    ],
  };
  const fetchJson = async (path) => (path === LISTS_PATH ? { status: 200, json: lists } : { status: 200, json: { Id: GUID_WEB, ServerRelativeUrl: '/sites/demo' } });
  const { resolved } = await resolvePlaceholders(findPlaceholders('{{ListId:Archive}}'), fetchJson);
  assert.match(resolved[0].error, /^Ambiguous/);
});

test('URL placeholders: values, any case, root site, unknown names left alone', () => {
  const v = urlPlaceholderValues({ webUrl: 'https://contoso.sharepoint.com/sites/demo/sub/', siteCollectionUrl: 'https://contoso.sharepoint.com/sites/demo' });
  assert.deepEqual(v, {
    WebUrl: 'https://contoso.sharepoint.com/sites/demo/sub',
    WebServerRelativeUrl: '/sites/demo/sub',
    SiteUrl: 'https://contoso.sharepoint.com/sites/demo',
    SiteServerRelativeUrl: '/sites/demo',
  });
  const root = urlPlaceholderValues({ webUrl: 'https://contoso.sharepoint.com', siteCollectionUrl: 'https://contoso.sharepoint.com' });
  assert.equal(root.WebServerRelativeUrl, '');
  const text = "POST {{WebUrl}}/_api/web/getList('{{ webserverrelativeurl }}/Lists/x') {{Other}}";
  assert.deepEqual(findUrlPlaceholders(text), ['WebUrl', 'WebServerRelativeUrl']);
  assert.equal(applyUrlPlaceholders(text, v), "POST https://contoso.sharepoint.com/sites/demo/sub/_api/web/getList('/sites/demo/sub/Lists/x') {{Other}}");
});

test('findCreator finds the list creation by title or last URL segment, only before a position', () => {
  const ops = parseBatch(readSample('Batch1.txt')).operations; // creates crRequests ... crProductionSites
  assert.equal(findCreator(ops, '/sites/rd-hub-projects/Lists/crDepartments').index, 2);
  assert.equal(findCreator(ops, 'Lists/crAttachments').index, 4);
  assert.equal(findCreator(ops, 'crattachments').index, 4);
  assert.equal(findCreator(ops, 'crDepartments', 2), null);
  assert.equal(findCreator(ops, 'nope'), null);
});

test('no list placeholders: the lists are not read', async () => {
  const fetchJson = sharePoint();
  const { listGuids } = await resolvePlaceholders(findPlaceholders('{{WebId}}'), fetchJson);
  assert.deepEqual(fetchJson.calls, [WEB_PATH]);
  assert.equal(listGuids, null);
});

test('unresolved placeholders are reported, not replaced', async () => {
  const { resolved } = await resolvePlaceholders(findPlaceholders('{{ListId:Missing}} {{ListId:/sites/demo/Lists/nope}}'), sharePoint());
  assert.deepEqual(resolved.map((r) => r.error), ['List not found.', 'List not found.']);
  assert.equal(applyPlaceholders('{{ListId:Missing}}', resolved), '{{ListId:Missing}}');

  const denied = async () => ({ status: 403, json: null, message: 'Access denied.' });
  const r2 = await resolvePlaceholders(findPlaceholders('{{ListId:A}} {{WebId}}'), denied);
  assert.deepEqual(r2.resolved.map((r) => r.error), ['Reading the lists failed: HTTP 403: Access denied.', 'HTTP 403: Access denied.']);

  const down = async () => {
    throw new Error('Network down');
  };
  const r3 = await resolvePlaceholders(findPlaceholders('{{ListId:A}}'), down);
  assert.equal(r3.resolved[0].error, 'Reading the lists failed: Network down');
});

test('samples: lookup columns use the Power Automate alias and all of them are found', () => {
  for (const name of sampleNames()) {
    const text = readSample(name);
    const expected = new Set(
      [...text.matchAll(/objListGuids'\)\?\['([^']+)'\]/g)].map((m) => m[1].toLowerCase())
    );
    const found = findPlaceholders(text);
    assert.deepEqual(
      new Set(found.map((p) => p.ref.toLowerCase())),
      expected,
      name
    );
  }
  const crLists = findPlaceholders(readSample('Batch2.txt')).map((p) => p.ref);
  assert.ok(crLists.includes('/sites/rd-hub-projects/Lists/crRequests'));
});
