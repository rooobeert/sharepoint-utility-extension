import { test } from 'node:test';
import assert from 'node:assert/strict';
import { groupRoles, limitedOnly, scanGroupUsage } from '../extension/lib/group-usage.js';
import { createPermissionClient } from '../extension/lib/permission-client.js';

const ROOT = 'https://contoso.sharepoint.com/sites/demo';
const ra = (principalId, ...names) => ({ PrincipalId: principalId, RoleDefinitionBindings: names.map((n) => ({ Name: n, Hidden: n === 'Limited Access' })) });

/** A fake site collection: url -> { web, subs, lists, ras: { '': web assignments, [listId]: list assignments } } */
function fake(tree, fail = {}) {
  const calls = [];
  const get = async (url, path) => {
    calls.push(url + path);
    if (fail[url + path]) throw new Error(fail[url + path]);
    const n = tree[url];
    if (!n) throw new Error('HTTP 404');
    if (path.startsWith('/_api/web?')) return n.web;
    if (path.startsWith('/_api/web/webs')) return { value: (n.subs || []).map((Url) => ({ Url })) };
    if (path.startsWith('/_api/web/lists?')) return { value: n.lists || [] };
    if (path.startsWith('/_api/web/roleassignments')) return { value: n.ras[''] || [] };
    const m = path.match(/lists\(guid'([^']+)'\)\/roleassignments/);
    if (m) return { value: n.ras[m[1]] || [] };
    throw new Error(`unexpected ${path}`);
  };
  return { get, calls };
}

const TREE = {
  [ROOT]: {
    web: { Title: 'Root', HasUniqueRoleAssignments: true },
    subs: [`${ROOT}/hr`, `${ROOT}/inherit`],
    lists: [
      { Id: 'l1', Title: 'Docs', BaseType: 1, HasUniqueRoleAssignments: true, RootFolder: { ServerRelativeUrl: '/sites/demo/Docs' } },
      { Id: 'l2', Title: 'Tasks', BaseType: 0, HasUniqueRoleAssignments: false },
      { Id: 'l3', Title: 'Other', BaseType: 0, HasUniqueRoleAssignments: true },
    ],
    ras: { '': [ra(7, 'Contribute'), ra(8, 'Read')], l1: [ra(7, 'Read'), ra(9, 'Edit')], l3: [ra(9, 'Edit')] },
  },
  [`${ROOT}/hr`]: {
    web: { Title: 'HR', HasUniqueRoleAssignments: true },
    lists: [{ Id: 'h1', Title: 'Staff', BaseType: 0, HasUniqueRoleAssignments: true, RootFolder: { ServerRelativeUrl: '/sites/demo/hr/Lists/Staff' } }],
    ras: { '': [ra(7, 'Limited Access')], h1: [ra(7, 'Full Control')] },
  },
  [`${ROOT}/inherit`]: { web: { Title: 'Inherit', HasUniqueRoleAssignments: false }, lists: [], ras: {} },
};

test('groupRoles: only the roles of that principal', () => {
  const json = { value: [ra(7, 'Read', 'Edit'), ra(8, 'Full Control')] };
  assert.deepEqual(groupRoles(json, 7).map((r) => r.name), ['Read', 'Edit']);
  assert.deepEqual(groupRoles(json, 99), []);
  assert.deepEqual(groupRoles(null, 7), []);
});

test('limitedOnly', () => {
  assert.equal(limitedOnly([]), false);
  assert.equal(limitedOnly([{ name: 'Limited Access', hidden: true }]), true);
  assert.equal(limitedOnly([{ name: 'Read', hidden: false }, { name: 'Limited Access', hidden: true }]), false);
});

test('scan: finds the group on webs, lists and libraries of all subsites', async () => {
  const { get } = fake(TREE);
  const res = await scanGroupUsage({ get }, ROOT, 7);
  assert.equal(res.webs, 3);
  assert.equal(res.lists, 4);
  assert.deepEqual(res.errors, []);
  const summary = res.places.map((p) => [p.type, p.title, p.roles.join('+'), p.limited]);
  assert.deepEqual(summary, [
    ['web', 'Root', 'Contribute', false],
    ['library', 'Docs', 'Read', false],
    ['web', 'HR', 'Limited Access', true],
    ['list', 'Staff', 'Full Control', false],
  ]);
  assert.equal(res.places[1].url, 'https://contoso.sharepoint.com/sites/demo/Docs');
  assert.equal(res.places[3].webTitle, 'HR');
});

test('scan: objects that inherit are not asked for their assignments', async () => {
  const { get, calls } = fake(TREE);
  await scanGroupUsage({ get }, ROOT, 7);
  assert.ok(!calls.some((c) => c.startsWith(`${ROOT}/inherit/_api/web/roleassignments`)));
  assert.ok(!calls.some((c) => c.includes("guid'l2'")));
});

test('scan: a group that is nowhere gives no places', async () => {
  const { get } = fake(TREE);
  const res = await scanGroupUsage({ get }, ROOT, 12345);
  assert.deepEqual(res.places, []);
});

test('scan: a part that cannot be read is reported and the crawl goes on', async () => {
  const { get } = fake(TREE, { [`${ROOT}/hr/_api/web/lists(guid'h1')/roleassignments?$select=PrincipalId&$expand=RoleDefinitionBindings&$top=5000`]: 'HTTP 403: no access' });
  const res = await scanGroupUsage({ get }, ROOT, 7);
  assert.equal(res.errors.length, 1);
  assert.match(res.errors[0].where, /HR \/ Staff/);
  assert.match(res.errors[0].message, /403/);
  assert.equal(res.places.length, 3);
});

test('scan: a subsite that cannot be opened is reported', async () => {
  const tree = { ...TREE, [ROOT]: { ...TREE[ROOT], subs: [`${ROOT}/secret`] } };
  const res = await scanGroupUsage({ get: fake(tree).get }, ROOT, 7);
  assert.equal(res.webs, 1);
  assert.deepEqual(res.errors.map((e) => e.where), [`${ROOT}/secret`]);
});

test('scan: stops when asked and says so', async () => {
  let n = 0;
  const res = await scanGroupUsage({ get: fake(TREE).get, shouldStop: () => n++ >= 1 }, ROOT, 7);
  assert.equal(res.stopped, true);
  assert.equal(res.webs, 1);
});

test('scan: reports progress per web', async () => {
  const events = [];
  await scanGroupUsage({ get: fake(TREE).get, onEvent: (e) => events.push(e.url) }, ROOT, 7);
  assert.deepEqual(events, [ROOT, `${ROOT}/hr`, `${ROOT}/inherit`]);
});

test('client: starts at the site collection root, not the web in the tab', async () => {
  const f = fake(TREE);
  const request = async (req) => {
    const path = req.url;
    let body;
    if (path.endsWith('/_api/site?$select=Url')) body = { Url: ROOT };
    else {
      const url = Object.keys(TREE).sort((a, b) => b.length - a.length).find((u) => path.startsWith(u + '/_api'));
      body = await f.get(url, path.slice(url.length));
    }
    return { status: 200, text: JSON.stringify(body), headers: {} };
  };
  const client = createPermissionClient({ request });
  const res = await client.findGroupUsage(`${ROOT}/hr`, 7);
  assert.equal(res.webs, 3);
});

function entraClient(user) {
  const f = fake(TREE);
  const seen = [];
  const request = async (req) => {
    seen.push(req.url);
    if (req.url.endsWith('/_api/site?$select=Url')) return { status: 200, text: JSON.stringify({ Url: ROOT }), headers: {} };
    if (req.url.includes('/siteusers/getbyloginname')) {
      return user ? { status: 200, text: JSON.stringify({ Id: user }), headers: {} } : { status: 404, text: '{}', headers: {} };
    }
    const url = Object.keys(TREE).sort((a, b) => b.length - a.length).find((u) => req.url.startsWith(u + '/_api'));
    return { status: 200, text: JSON.stringify(await f.get(url, req.url.slice(url.length))), headers: {} };
  };
  return { client: createPermissionClient({ request }), seen };
}

test('client: an Entra group is looked up by login name in the site users, then crawled', async () => {
  const { client, seen } = entraClient(7);
  const login = "c:0t.c|tenant|abc-123";
  const res = await client.findGroupUsage(ROOT, { login });
  assert.equal(res.places.length, 4);
  const lookup = seen.find((u) => u.includes('getbyloginname'));
  assert.ok(lookup.startsWith(`${ROOT}/_api/web/siteusers/getbyloginname(@v)?@v='c%3A0t.c%7Ctenant%7Cabc-123'`));
});

test('client: an Entra group that is not a site user has no permissions and nothing is crawled', async () => {
  const { client, seen } = entraClient(null);
  const res = await client.findGroupUsage(ROOT, { login: 'c:0t.c|tenant|zzz' });
  assert.equal(res.notInSite, true);
  assert.deepEqual(res.places, []);
  assert.ok(!seen.some((u) => u.includes('/_api/web/lists')));
});
