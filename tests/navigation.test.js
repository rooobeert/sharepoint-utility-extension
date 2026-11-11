import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  LINKLESS,
  defaultTarget,
  detectNavigation,
  listPrefix,
  groupListsByPrefix,
  editorUrl,
  sharePointUrl,
  isExternalUrl,
  linksToList,
  treeFromNodes,
  cloneTree,
  treeSignature,
  validateTree,
  planSave,
  planIsEmpty,
} from '../extension/lib/navigation.js';
import { createNavigator } from '../extension/lib/navigator.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const ORIGIN = 'https://contoso.sharepoint.com';
const list = (url, title, extra = {}) => ({ id: url, title, url: `/sites/demo/${url}`, template: 100, system: false, ...extra });

// ---------------------------------------------------------------- grouping

test('listPrefix takes the lowercase solution tag of the internal name', () => {
  assert.equal(listPrefix('/sites/demo/Lists/hrpLocations'), 'hrp');
  assert.equal(listPrefix('/sites/demo/Lists/tdasTermDeals'), 'tdas');
  assert.equal(listPrefix('/sites/demo/hrpDocuments'), 'hrp');
  assert.equal(listPrefix('/sites/demo/Lists/crm_Contacts'), 'crm');
  assert.equal(listPrefix('/sites/demo/Lists/abc2024'), 'abc');
  assert.equal(listPrefix('/sites/demo/Shared Documents'), null);
  assert.equal(listPrefix('/sites/demo/SitePages'), null);
  assert.equal(listPrefix('/sites/demo/Lists/tasks'), null);
  assert.equal(listPrefix('/sites/demo/Lists/aLongList'), null); // one letter is no tag
});

test('groupListsByPrefix groups by tag, needs two lists, skips system lists', () => {
  const groups = groupListsByPrefix([
    list('Lists/hrpLocations', 'Locations'),
    list('Lists/hrpEmployees', 'Employees'),
    list('Lists/tdasTermDeals', 'Term deals'),
    list('Lists/tdasClients', 'Clients'),
    list('Lists/myList', 'Mine'),
    list('Shared Documents', 'Documents', { template: 101 }),
    list('_catalogs/masterpage', 'Master Page Gallery', { system: true }),
  ]);
  assert.deepEqual(groups.map((g) => [g.label, g.lists.map((l) => l.title)]), [
    ['HRP', ['Employees', 'Locations']],
    ['TDAS', ['Clients', 'Term deals']],
    ['Other lists and libraries', ['Documents', 'Mine']],
  ]);
  assert.equal(groups[2].prefix, null);
});

test('defaultTarget: communication sites use the top navigation', () => {
  assert.equal(defaultTarget({ WebTemplate: 'SITEPAGEPUBLISHING' }), 'topnav');
  assert.equal(defaultTarget({ WebTemplate: 'GROUP' }), 'quicklaunch');
  assert.equal(defaultTarget({ WebTemplate: 'STS' }), 'quicklaunch');
  assert.equal(defaultTarget(null), 'quicklaunch');
});

test('detectNavigation reads the layout from the site settings', () => {
  const d = (web) => [detectNavigation(web).target, detectNavigation(web).layout];
  assert.deepEqual(d({ WebTemplate: 'SITEPAGEPUBLISHING', MegaMenuEnabled: true }), ['topnav', 'top']);
  assert.match(detectNavigation({ WebTemplate: 'SITEPAGEPUBLISHING', MegaMenuEnabled: true }).text, /mega menu/);
  assert.deepEqual(d({ WebTemplate: 'GROUP', QuickLaunchEnabled: true, HorizontalQuickLaunch: false }), ['quicklaunch', 'vertical']);
  assert.deepEqual(d({ WebTemplate: 'GROUP', QuickLaunchEnabled: true, HorizontalQuickLaunch: true }), ['quicklaunch', 'horizontal']);
  assert.deepEqual(d({ WebTemplate: 'STS', QuickLaunchEnabled: false }), ['quicklaunch', 'off']);
});

// ---------------------------------------------------------------- URLs

test('headings: empty URL in the editor, linkless header in SharePoint', () => {
  assert.equal(editorUrl(LINKLESS), '');
  assert.equal(editorUrl('http://linkless.header'), '');
  assert.equal(editorUrl('  /sites/demo  '), '/sites/demo');
  assert.equal(sharePointUrl(''), LINKLESS);
  assert.equal(sharePointUrl('/sites/demo/Lists/x'), '/sites/demo/Lists/x');
  assert.equal(isExternalUrl(''), true);
  assert.equal(isExternalUrl('https://example.com'), true);
  assert.equal(isExternalUrl('/sites/demo'), false);
});

test('linksToList matches the root folder, views below it and full URLs', () => {
  const url = '/sites/demo/Lists/hrpLocations';
  assert.ok(linksToList('/sites/demo/Lists/hrpLocations', url, ORIGIN));
  assert.ok(linksToList('/sites/demo/Lists/hrpLocations/AllItems.aspx', url, ORIGIN));
  assert.ok(linksToList('https://contoso.sharepoint.com/sites/demo/lists/HRPLOCATIONS/', url, ORIGIN));
  assert.ok(linksToList('/sites/demo/Shared%20Documents/Forms/AllItems.aspx', '/sites/demo/Shared Documents', ORIGIN));
  assert.ok(!linksToList('/sites/demo/Lists/hrpLocationsOld', url, ORIGIN));
  assert.ok(!linksToList('https://other.sharepoint.com/sites/demo/Lists/hrpLocations', url, ORIGIN));
  assert.ok(!linksToList(LINKLESS, url, ORIGIN));
});

// ---------------------------------------------------------------- tree

const REST = [
  { Id: 1, Title: 'Home', Url: '/sites/demo', Children: [] },
  {
    Id: 2,
    Title: 'HRP',
    Url: LINKLESS,
    Children: [
      { Id: 3, Title: 'Locations', Url: '/sites/demo/Lists/hrpLocations', Children: [] },
      { Id: 4, Title: 'Employees', Url: '/sites/demo/Lists/hrpEmployees', Children: [] },
    ],
  },
  { Id: 5, Title: 'Documents', Url: '/sites/demo/Shared Documents', Children: [] },
];
const original = () => treeFromNodes(REST);
const node = (title, url = '', children = []) => ({ key: `k-${title}`, id: null, title, url, children });

test('treeFromNodes and signature', () => {
  const t = original();
  assert.equal(t[1].url, '');
  assert.equal(t[1].children[0].key, 'id3');
  const copy = cloneTree(t);
  assert.equal(treeSignature(copy), treeSignature(t));
  copy[1].children[0].title = 'Standorte';
  assert.notEqual(treeSignature(copy), treeSignature(t));
  assert.equal(t[1].children[0].title, 'Locations'); // deep copy
});

test('validateTree reports empty titles, bad URLs and a fourth level', () => {
  const deep = [node('A', '', [node('B', '', [node('C', '', [node('D', '/x')])])])];
  const problems = validateTree([node('', '/x'), node('Bad', 'Lists/x'), ...deep]);
  assert.equal(problems.length, 3);
  assert.match(problems[0], /no title/);
  assert.match(problems[1], /must start with/);
  assert.match(problems[2], /level 4/);
  assert.deepEqual(validateTree(original()), []);
});

// ---------------------------------------------------------------- plan

test('planSave: no changes', () => {
  assert.ok(planIsEmpty(planSave(original(), original())));
});

test('planSave: rename, relink and append keep existing nodes', () => {
  const t = original();
  t[1].children[0].title = 'Standorte';
  t[2].url = '/sites/demo/Shared Documents/Forms/AllItems.aspx';
  t[1].children.push(node('Absences', '/sites/demo/Lists/hrpAbsences'));
  t.push(node('TDAS', '', [node('Deals', '/sites/demo/Lists/tdasDeals')]));
  const plan = planSave(original(), t);
  assert.deepEqual(plan.deletes, []);
  assert.deepEqual(plan.updates.map((u) => [u.id, u.title, u.url]), [
    [3, 'Standorte', '/sites/demo/Lists/hrpLocations'],
    [5, 'Documents', '/sites/demo/Shared Documents/Forms/AllItems.aspx'],
  ]);
  assert.deepEqual(plan.creates.map((c) => [c.title, c.parent, c.external, c.recreated]), [
    ['Absences', { id: 2 }, false, false],
    ['TDAS', null, true, false],
    ['Deals', { key: 'k-TDAS' }, false, false],
  ]);
});

test('planSave: removing a heading deletes only the top-most node', () => {
  const t = original();
  t.splice(1, 1);
  const plan = planSave(original(), t);
  assert.deepEqual(plan.deletes, [{ id: 2, title: 'HRP' }]);
  assert.deepEqual(plan.removed.map((r) => r.title), ['HRP', 'Locations', 'Employees']);
  assert.deepEqual(plan.creates, []);
});

test('planSave: a reorder recreates the moved node and everything after it on its level', () => {
  const t = original();
  t.unshift(t.pop()); // Documents first
  const plan = planSave(original(), t);
  // Documents is kept (first), Home and HRP come after it in a different order -> recreated
  assert.deepEqual(plan.deletes.map((d) => d.id), [1, 2]);
  assert.deepEqual(plan.creates.map((c) => [c.title, c.recreated]), [
    ['Home', true],
    ['HRP', true],
    ['Locations', true],
    ['Employees', true],
  ]);
  assert.deepEqual(plan.removed, []); // all recreated, nothing lost
});

test('planSave: a link moved into another heading is recreated there', () => {
  const t = original();
  const docs = t.pop();
  t[1].children.push(docs);
  const plan = planSave(original(), t);
  assert.deepEqual(plan.deletes, [{ id: 5, title: 'Documents' }]);
  assert.deepEqual(plan.creates.map((c) => [c.title, c.parent, c.recreated]), [['Documents', { id: 2 }, true]]);
  assert.deepEqual(plan.removed, []);
});

test('planSave: a new node before existing ones forces those to be recreated', () => {
  const t = original();
  t[1].children.unshift(node('Absences', '/sites/demo/Lists/hrpAbsences'));
  const plan = planSave(original(), t);
  assert.deepEqual(plan.deletes.map((d) => d.id), [3, 4]);
  assert.deepEqual(plan.creates.map((c) => c.title), ['Absences', 'Locations', 'Employees']);
});

// ---------------------------------------------------------------- navigator

const reply = (status, body) => ({ phase: 'request', status, statusText: '', headers: {}, text: body == null ? '' : JSON.stringify(body) });

function fakeNav(handler) {
  const sent = [];
  const request = async (req) => {
    sent.push(req);
    return handler(req, sent.length);
  };
  return { sent, request };
}

test('navigator.load reads three levels', async () => {
  const { sent, request } = fakeNav((req) => {
    const path = req.url.slice(SITE.length);
    if (path === '/_api/web/navigation/QuickLaunch') return reply(200, { value: [{ Id: 1, Title: 'A', Url: LINKLESS }] });
    if (path === '/_api/web/navigation/GetNodeById(1)/Children') return reply(200, { value: [{ Id: 2, Title: 'B', Url: '/b' }] });
    if (path === '/_api/web/navigation/GetNodeById(2)/Children') return reply(200, { value: [{ Id: 3, Title: 'C', Url: '/c' }] });
    return reply(404, { 'odata.error': { message: { value: 'nope' } } });
  });
  const nodes = await createNavigator({ request }).load(SITE, 'quicklaunch');
  assert.equal(nodes[0].Children[0].Children[0].Title, 'C');
  assert.deepEqual(nodes[0].Children[0].Children[0].Children, []); // level 3 is not expanded
  assert.equal(sent.length, 3);
  assert.ok(sent.every((r) => r.method === 'GET' && !r.digestFrom));
});

test('navigator.save deletes, updates, creates in order and wires new parents', async () => {
  let nextId = 100;
  const { sent, request } = fakeNav((req) => {
    if (req.headers['X-HTTP-Method']) return reply(204);
    return reply(201, { Id: nextId++ });
  });
  const plan = {
    deletes: [{ id: 5, title: 'Old' }],
    updates: [{ id: 3, title: 'Standorte', url: '/sites/demo/Lists/hrpLocations' }],
    creates: [
      { key: 'h', parent: null, title: 'TDAS', url: '', external: true },
      { key: 'l', parent: { key: 'h' }, title: 'Deals', url: '/sites/demo/Lists/tdasDeals', external: false },
      { key: 'm', parent: { id: 2 }, title: 'Absences', url: '/sites/demo/Lists/hrpAbsences', external: false },
    ],
  };
  const results = await createNavigator({ request }).save(SITE, 'topnav', plan);
  assert.ok(results.every((r) => r.ok));
  const calls = sent.map((r) => [r.headers['X-HTTP-Method'] || r.method, r.url.slice(SITE.length)]);
  assert.deepEqual(calls, [
    ['DELETE', '/_api/web/navigation/GetNodeById(5)'],
    ['MERGE', '/_api/web/navigation/GetNodeById(3)'],
    ['POST', '/_api/web/navigation/TopNavigationBar'],
    ['POST', '/_api/web/navigation/GetNodeById(100)/Children'],
    ['POST', '/_api/web/navigation/GetNodeById(2)/Children'],
  ]);
  assert.ok(sent.every((r) => r.digestFrom === SITE));
  const heading = JSON.parse(sent[2].body);
  assert.deepEqual(heading, { __metadata: { type: 'SP.NavigationNode' }, Title: 'TDAS', Url: LINKLESS, IsExternal: true });
  assert.equal(JSON.parse(sent[3].body).IsExternal, undefined);
});

test('navigator.save goes on after a failure and skips children of a failed node', async () => {
  const { sent, request } = fakeNav((req) => {
    if (req.headers['X-HTTP-Method'] === 'DELETE') return reply(404, { 'odata.error': { message: { value: 'gone' } } });
    if (req.url.endsWith('/QuickLaunch')) return reply(400, { 'odata.error': { message: { value: 'Bad URL' } } });
    return reply(201, { Id: 9 });
  });
  const plan = {
    deletes: [{ id: 5, title: 'Old' }],
    updates: [],
    creates: [
      { key: 'h', parent: null, title: 'TDAS', url: '', external: true },
      { key: 'l', parent: { key: 'h' }, title: 'Deals', url: '/x', external: false },
      { key: 'm', parent: { id: 2 }, title: 'Absences', url: '/y', external: false },
    ],
  };
  const results = await createNavigator({ request }).save(SITE, 'quicklaunch', plan);
  assert.deepEqual(results.map((r) => [r.title, r.ok]), [['Old', true], ['TDAS', false], ['Deals', false], ['Absences', true]]);
  assert.match(results[1].error, /HTTP 400: Bad URL/);
  assert.match(results[2].error, /parent could not be created/);
  assert.equal(sent.length, 3); // Deals was never sent
});

test('navigator.settings falls back when SharePoint rejects the $select', async () => {
  const { sent, request } = fakeNav((req, n) =>
    n === 1 ? reply(400, { 'odata.error': { message: { value: 'property does not exist' } } }) : reply(200, { WebTemplate: 'GROUP', QuickLaunchEnabled: true }));
  const web = await createNavigator({ request }).settings(SITE);
  assert.equal(web.WebTemplate, 'GROUP');
  assert.match(sent[0].url, /HorizontalQuickLaunch,MegaMenuEnabled$/);
  assert.match(sent[1].url, /\$select=WebTemplate,QuickLaunchEnabled$/);
});

test('navigator retries on throttling', async () => {
  const waits = [];
  const { request } = fakeNav((req, n) => (n === 1 ? { ...reply(429), headers: { 'retry-after': '2' } } : reply(200, { value: [] })));
  const nodes = await createNavigator({ request, sleep: async (ms) => waits.push(ms) }).load(SITE, 'quicklaunch');
  assert.deepEqual(nodes, []);
  assert.deepEqual(waits, [2000]);
});

// ---------------------------------------------------------------- copy and paste

import { portableUrl, resolvePortableUrl, exportNavigation, parseNavigationExport } from '../extension/lib/navigation.js';

const SRC = { origin: 'https://t.sharepoint.com', webServerRelativeUrl: '/sites/dev' };
const DST = { site: 'https://t.sharepoint.com/sites/prod', webServerRelativeUrl: '/sites/prod' };

test('portableUrl: links into the web follow the site; others stay', () => {
  assert.equal(portableUrl('https://t.sharepoint.com/sites/dev/Lists/A/AllItems.aspx', SRC), '{{WebServerRelativeUrl}}/Lists/A/AllItems.aspx');
  assert.equal(portableUrl('/sites/dev/Lists/A', SRC), '{{WebServerRelativeUrl}}/Lists/A');
  assert.equal(portableUrl('/sites/dev', SRC), '{{WebServerRelativeUrl}}');
  assert.equal(portableUrl('/sites/devtwo/Lists/A', SRC), '/sites/devtwo/Lists/A');
  assert.equal(portableUrl('https://t.sharepoint.com/sites/other/x', SRC), 'https://t.sharepoint.com/sites/other/x');
  assert.equal(portableUrl('https://example.com/x', SRC), 'https://example.com/x');
  assert.equal(portableUrl(LINKLESS, SRC), '');
  assert.equal(portableUrl('/Lists/A', { origin: SRC.origin, webServerRelativeUrl: '' }), '{{WebServerRelativeUrl}}/Lists/A');
});

test('resolvePortableUrl: placeholders become the target web; a root web link is /', () => {
  assert.equal(resolvePortableUrl('{{WebServerRelativeUrl}}/Lists/A', DST), '/sites/prod/Lists/A');
  assert.equal(resolvePortableUrl('{{WebServerRelativeUrl}}', DST), '/sites/prod');
  assert.equal(resolvePortableUrl('{{WebServerRelativeUrl}}', { site: 'https://t.sharepoint.com', webServerRelativeUrl: '' }), '/');
  assert.equal(resolvePortableUrl('{{WebUrl}}/SitePages/Home.aspx', DST), 'https://t.sharepoint.com/sites/prod/SitePages/Home.aspx');
  assert.equal(resolvePortableUrl('https://example.com/x', DST), 'https://example.com/x');
});

test('export then paste on another site', () => {
  const tree = [
    { key: 'a', id: 1, title: 'HR', url: '', children: [
      { key: 'b', id: 2, title: 'Requests', url: '/sites/dev/Lists/hrRequests', children: [] },
      { key: 'c', id: 3, title: 'Docs', url: 'https://example.com/d', children: [] },
    ] },
  ];
  const text = exportNavigation(tree, { target: 'quicklaunch', site: 'https://t.sharepoint.com/sites/dev', ...SRC });
  const out = parseNavigationExport(text, DST);
  assert.equal(out.count, 3);
  assert.equal(out.target, 'quicklaunch');
  assert.equal(out.nodes[0].title, 'HR');
  assert.equal(out.nodes[0].url, '');
  assert.equal(out.nodes[0].children[0].url, '/sites/prod/Lists/hrRequests');
  assert.equal(out.nodes[0].children[1].url, 'https://example.com/d');
  assert.ok(out.nodes[0].id === null && out.nodes[0].key.startsWith('new'));
});

test('parseNavigationExport: bare arrays work, bad input gives readable errors', () => {
  assert.equal(parseNavigationExport('[{"title":"A","url":"/x"}]', DST).nodes[0].url, '/x');
  assert.throws(() => parseNavigationExport('nope', DST), /not valid JSON/);
  assert.throws(() => parseNavigationExport('{"foo":1}', DST), /not a navigation copied/);
  assert.throws(() => parseNavigationExport('{"format":"sp-utility-navigation","nodes":[{"url":"/x"}]}', DST), /Item 1 has no title/);
  assert.throws(() => parseNavigationExport('{"format":"sp-utility-navigation","nodes":[{"title":"A","children":[{"title":5}]}]}', DST), /Item 1\.1/);
});

// ---------------------------------------------------------------- backups

import { backupKey, snapshotTree, countNodes, addBackup, treeFromSnapshot, MAX_BACKUPS_PER_NAV } from '../extension/lib/navigation.js';

test('backups: key per site and navigation, snapshot and rebuild', () => {
  assert.equal(backupKey('https://T.sharepoint.com/sites/a/', 'quicklaunch'), 'https://t.sharepoint.com/sites/a|quicklaunch');
  const tree = [{ key: 'id1', id: 1, title: 'HR', url: LINKLESS, children: [{ key: 'id2', id: 2, title: 'R', url: '/x', children: [] }] }];
  const snap = snapshotTree(tree);
  assert.deepEqual(snap, [{ title: 'HR', url: '', children: [{ title: 'R', url: '/x', children: [] }] }]);
  assert.equal(countNodes(snap), 2);
  const back = treeFromSnapshot(snap);
  assert.ok(back[0].id === null && back[0].children[0].id === null && back[0].key !== back[0].children[0].key);
  assert.equal(back[0].children[0].url, '/x');
});

test('addBackup keeps the newest few per navigation and limits the number of sites', () => {
  let all = {};
  for (let i = 0; i < MAX_BACKUPS_PER_NAV + 2; i++) all = addBackup(all, 'a|quicklaunch', { at: `2026-01-0${i + 1}`, tree: [] });
  assert.equal(all['a|quicklaunch'].length, MAX_BACKUPS_PER_NAV);
  assert.equal(all['a|quicklaunch'][0].at, '2026-01-07');
  for (let i = 0; i < 40; i++) all = addBackup(all, `s${i}|topnav`, { at: `2026-02-${String(i + 1).padStart(2, '0')}`, tree: [] });
  assert.ok(Object.keys(all).length <= 30);
  assert.ok(all['s39|topnav']);
  assert.ok(!all['s0|topnav']);
});

// ---------------------------------------------------------------- hub navigation

import { parseHubSite } from '../extension/lib/navigation.js';

test('parseHubSite reads HubSiteData in its string and object forms', () => {
  const inner = { name: 'Intranet', url: 'https://t.sharepoint.com/sites/intranet/', navigation: [] };
  assert.deepEqual(parseHubSite({ value: JSON.stringify(inner) }), { url: 'https://t.sharepoint.com/sites/intranet', name: 'Intranet' });
  assert.deepEqual(parseHubSite(inner), { url: 'https://t.sharepoint.com/sites/intranet', name: 'Intranet' });
  assert.equal(parseHubSite({ value: 'null' }), null);
  assert.equal(parseHubSite({ value: '' }), null);
  assert.equal(parseHubSite(null), null);
  assert.equal(parseHubSite({ value: 'not json' }), null);
  assert.equal(parseHubSite({ url: 'http://x/y' }), null);
});

import { hubFromSite } from '../extension/lib/navigation.js';

test('hubFromSite: hub site itself, associated site, no hub', () => {
  assert.deepEqual(hubFromSite({ IsHubSite: true, Url: 'https://t.sharepoint.com/sites/hub/' }, null), { url: 'https://t.sharepoint.com/sites/hub', name: '' });
  const hubs = { value: [{ SiteId: 'AAAA0000-0000-0000-0000-000000000001', SiteUrl: 'https://t.sharepoint.com/sites/intranet', Title: 'Intranet' }] };
  assert.deepEqual(hubFromSite({ HubSiteId: '{aaaa0000-0000-0000-0000-000000000001}' }, hubs), { url: 'https://t.sharepoint.com/sites/intranet', name: 'Intranet' });
  assert.equal(hubFromSite({ HubSiteId: '00000000-0000-0000-0000-000000000000' }, hubs), null);
  assert.equal(hubFromSite({ HubSiteId: 'bbbb0000-0000-0000-0000-000000000002' }, hubs), null);
});
