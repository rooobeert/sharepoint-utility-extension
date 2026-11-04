import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  pickerPrincipals,
  pickerQuery,
  loginKind,
  roleLevels,
  defaultRoleId,
  siteGroups,
  defaultGroupIds,
  assignments,
  addEntries,
  planList,
  countActions,
  listOperations,
  packBatches,
  buildPermissionBatch,
  hasChanges,
  isLibrary,
  itemLevelText,
} from '../extension/lib/permissions.js';
import { createPermissionClient } from '../extension/lib/permission-client.js';
import { parseBatch } from '../extension/lib/batch.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const LIST = { id: 'aaaa-1', title: 'Contracts' };

const ROLES = roleLevels({
  value: [
    { Id: 1073741829, Name: 'Full Control', RoleTypeKind: 5, Order: 1 },
    { Id: 1073741827, Name: 'Edit', RoleTypeKind: 6, Order: 48 },
    { Id: 1073741826, Name: 'Read', RoleTypeKind: 2, Order: 128 },
    { Id: 1073741825, Name: 'Limited Access', RoleTypeKind: 1, Order: 160, Hidden: true },
  ],
});
const FULL = 1073741829;
const EDIT = 1073741827;
const READ = 1073741826;
const LIMITED = 1073741825;
const WEB = { AssociatedOwnerGroup: { Id: 3 }, AssociatedMemberGroup: { Id: 5 }, AssociatedVisitorGroup: { Id: 4 } };
const SITE_DATA = { roles: ROLES, defaults: defaultGroupIds(WEB) };

const binding = (id) => ({ Id: id, Name: ROLES.find((r) => r.id === id)?.name || 'Limited Access', RoleTypeKind: ROLES.find((r) => r.id === id)?.kind ?? 1, Hidden: id === LIMITED });
const ra = (id, title, login, roleIds, type = 8) => ({ PrincipalId: id, Member: { Id: id, Title: title, LoginName: login, PrincipalType: type }, RoleDefinitionBindings: roleIds.map(binding) });
const inherited = assignments({
  value: [
    ra(3, 'Demo Owners', 'Demo Owners', [FULL]),
    ra(5, 'Demo Members', 'Demo Members', [EDIT]),
    ra(4, 'Demo Visitors', 'Demo Visitors', [READ]),
    ra(12, 'Project Group', 'Project Group', [EDIT]),
    ra(20, 'Anna', 'i:0#.f|membership|anna@contoso.com', [LIMITED], 1),
  ],
});

const anna = { key: 'l:i:0#.f|membership|anna@contoso.com', id: null, login: 'i:0#.f|membership|anna@contoso.com', title: 'Anna', kind: 'User' };
const ben = { key: 'l:i:0#.f|membership|ben@contoso.com', id: null, login: 'i:0#.f|membership|ben@contoso.com', title: 'Ben', kind: 'User' };
const members = { key: 'g:5', id: 5, login: null, title: 'Demo Members', kind: 'SPO Group' };
const entry = (principal, roleId) => ({ listId: LIST.id, principal, roleId });
const types = (plan) => plan.actions.map((a) => `${a.type} ${a.title}${a.role ? ` ${a.role.name}` : ''}`);

// ---------------------------------------------------------------- principals and site data

test('pickerPrincipals: users and Entra groups from the JSON string; SharePoint groups left out', () => {
  const raw = JSON.stringify([
    { Key: 'i:0#.f|membership|anna@contoso.com', DisplayText: 'Anna', EntityType: 'User', EntityData: { Email: 'anna@contoso.com' } },
    { Key: 'c:0t.c|tenant|1234', DisplayText: 'Finance', EntityType: 'SecGroup', EntityData: {} },
    { Key: 'c:0o.c|federateddirectoryclaimprovider|5678', DisplayText: 'Project X', EntityType: 'SecGroup', EntityData: {} },
    { Key: 'Demo Members', DisplayText: 'Demo Members', EntityType: 'SPGroup', EntityData: { SPGroupID: '5' } },
    { Key: '', DisplayText: 'Unresolved', EntityType: 'User' },
  ]);
  const found = pickerPrincipals({ d: { ClientPeoplePickerSearchUser: raw } });
  assert.deepEqual(found.map((p) => [p.title, p.kind, p.detail, p.key]), [
    ['Anna', 'User', null, 'l:i:0#.f|membership|anna@contoso.com'],
    ['Finance', 'Entra Group', 'Security group', 'l:c:0t.c|tenant|1234'],
    ['Project X', 'Entra Group', 'Microsoft 365 group', 'l:c:0o.c|federateddirectoryclaimprovider|5678'],
  ]);
  assert.equal(found[0].email, 'anna@contoso.com');
  assert.deepEqual(pickerPrincipals({ value: 'not json' }), []);
  assert.equal(loginKind('c:0-.f|rolemanager|spo-grid-all-users/abc').kind, 'Everyone');
});

test('pickerQuery asks for users and Entra groups, not SharePoint groups', () => {
  const q = JSON.parse(pickerQuery('ann')).queryParams;
  assert.equal(q.QueryString, 'ann');
  assert.equal(q.PrincipalType, 7);
  assert.equal(q.MaximumEntitySuggestions, 50);
  assert.equal(q.PrincipalSource, 15);
});

test('roleLevels hides Limited Access and keeps SharePoint order; Read is the default', () => {
  assert.deepEqual(ROLES.map((r) => r.name), ['Full Control', 'Edit', 'Read']);
  assert.equal(defaultRoleId(ROLES), READ);
  assert.equal(defaultRoleId([]), null);
});

test('siteGroups: default groups first, flagged', () => {
  const groups = siteGroups({ value: [{ Id: 12, Title: 'Alpha' }, { Id: 4, Title: 'Demo Visitors' }, { Id: 3, Title: 'Demo Owners' }, { Id: 5, Title: 'Demo Members' }] }, WEB);
  assert.deepEqual(groups.map((g) => [g.title, g.isDefault]), [
    ['Demo Owners', true], ['Demo Members', true], ['Demo Visitors', true], ['Alpha', false],
  ]);
  assert.deepEqual(SITE_DATA.defaults, { ids: [3, 5, 4], owners: 3, members: 5, visitors: 4 });
});

test('addEntries builds lists x principals and skips duplicates', () => {
  const plan = new Map();
  const picks = [{ principal: anna, roleId: READ }, { principal: members, roleId: EDIT }];
  assert.equal(addEntries(plan, ['l1', 'l2'], picks), 4);
  assert.equal(addEntries(plan, ['l2', 'l3'], picks), 2);
  assert.equal(plan.size, 6);
});

// ---------------------------------------------------------------- planning

test('add mode on an inheriting list: break with copy, add, nothing removed', () => {
  const plan = planList(LIST, { unique: false, assignments: inherited }, [entry(ben, EDIT), entry(anna, READ)], { mode: 'add' }, SITE_DATA);
  assert.equal(plan.breaks, true);
  assert.deepEqual(types(plan), ['add Anna Read', 'add Ben Edit']);
  // Anna has an assignment (Limited Access): her id comes from it, no ensureuser needed.
  assert.equal(plan.actions[0].principal.id, 20);
  assert.equal(plan.actions[1].principal.id, null);
  assert.deepEqual(plan.warnings, []);
});

test('add mode skips a level the principal already has; nothing to do means no break', () => {
  const plan = planList(LIST, { unique: false, assignments: inherited }, [entry(members, EDIT)], { mode: 'add' }, SITE_DATA);
  assert.deepEqual(types(plan), ['skip Demo Members']);
  assert.equal(plan.breaks, false);
});

test('replace mode keeping default groups: others removed, Limited Access kept', () => {
  const plan = planList(LIST, { unique: false, assignments: inherited }, [entry(ben, READ)], { mode: 'replace', keepGroups: [3, 5, 4] }, SITE_DATA);
  assert.equal(plan.breaks, true);
  assert.deepEqual(types(plan), [
    'keep Demo Owners', 'keep Demo Members', 'keep Demo Visitors', 'remove Project Group Edit', 'keep Anna', 'add Ben Read',
  ]);
  assert.deepEqual(countActions(plan), { add: 1, remove: 1, keep: 4, skip: 0, set: 0 });
  assert.deepEqual(plan.warnings, []);
});

test('replace mode removing default groups warns about Owners and Full Control', () => {
  const plan = planList(LIST, { unique: true, assignments: inherited }, [entry(ben, EDIT)], { mode: 'replace', keepGroups: [] }, SITE_DATA);
  assert.equal(plan.breaks, false);
  assert.deepEqual(types(plan).filter((t) => t.startsWith('remove')), [
    'remove Demo Owners Full Control', 'remove Demo Members Edit', 'remove Demo Visitors Read', 'remove Project Group Edit',
  ]);
  assert.equal(plan.warnings.length, 2);
  assert.match(plan.warnings[0], /Owners group "Demo Owners"/);
  assert.match(plan.warnings[1], /nobody has Full Control/);
});

test('replace mode on a planned principal: other levels removed, planned one added', () => {
  const plan = planList(LIST, { unique: true, assignments: inherited }, [entry(members, READ)], { mode: 'replace', keepGroups: [3, 5, 4] }, SITE_DATA);
  assert.deepEqual(types(plan).filter((t) => !t.startsWith('keep')), ['remove Demo Members Edit', 'add Demo Members Read', 'remove Project Group Edit']);
});

// ---------------------------------------------------------------- batches

test('listOperations and buildPermissionBatch produce a batch the parser accepts', () => {
  const plan = planList(LIST, { unique: false, assignments: inherited }, [entry(ben, READ)], { mode: 'replace', keepGroups: [3, 5, 4] }, SITE_DATA);
  const ops = listOperations(SITE, plan, (p) => p.id ?? 31);
  assert.deepEqual(ops.map((o) => o.url.slice(SITE.length)), [
    "/_api/web/lists(guid'aaaa-1')/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)",
    "/_api/web/lists(guid'aaaa-1')/roleassignments/removeroleassignment(principalid=12,roledefid=1073741827)",
    "/_api/web/lists(guid'aaaa-1')/roleassignments/addroleassignment(principalid=31,roledefid=1073741826)",
  ]);
  const parsed = parseBatch(buildPermissionBatch(ops));
  assert.deepEqual(parsed.errors, []);
  assert.equal(parsed.operations.length, 3);
  assert.ok(parsed.operations.every((o) => o.method === 'POST' && o.changeset));
  assert.throws(() => listOperations(SITE, plan), /No principal id for "Ben"/);
});

test('packBatches keeps lists together and splits only oversized lists', () => {
  const ops = (n, tag) => Array.from({ length: n }, (_, i) => `${tag}${i}`);
  const b = packBatches([ops(60, 'a'), ops(30, 'b'), ops(20, 'c'), ops(0, 'd'), ops(150, 'e')], 100);
  assert.deepEqual(b.map((x) => x.length), [90, 20, 100, 50]);
  assert.equal(b[0][60], 'b0');
  assert.equal(b[1][0], 'c0');
  assert.equal(b[2][0], 'e0');
});

// ---------------------------------------------------------------- client

const json = (status, body) => ({ phase: 'request', status, statusText: '', headers: {}, text: JSON.stringify(body) });

test('apply: ensureuser for new principals, one $batch, results per list', async () => {
  const sent = [];
  const request = async (req) => {
    sent.push(req);
    if (req.url.endsWith('/_api/web/ensureuser')) {
      const login = JSON.parse(req.body).logonName;
      return login.includes('ben') ? json(200, { Id: 31 }) : json(404, { 'odata.error': { message: { value: 'User not found' } } });
    }
    if (req.url.endsWith('/_api/$batch')) {
      const n = (req.body.match(/^(POST|MERGE) /gm) || []).length;
      const part = '--resp\r\nContent-Type: application/http\r\n\r\nHTTP/1.1 200 OK\r\n\r\n{}\r\n';
      return { phase: 'request', status: 200, statusText: 'OK', headers: { 'content-type': 'multipart/mixed; boundary=resp' }, text: part.repeat(n) + '--resp--\r\n' };
    }
    throw new Error(`unexpected ${req.url}`);
  };
  const ghost = { ...ben, key: 'l:ghost', login: 'i:0#.f|membership|ghost@contoso.com', title: 'Ghost' };
  const plan = planList(LIST, { unique: false, assignments: inherited }, [entry(ben, READ), entry(ghost, READ)], { mode: 'add' }, SITE_DATA);
  const result = await createPermissionClient({ request, sleep: async () => {} }).apply(SITE, [plan]);

  assert.equal(sent.filter((r) => r.url.endsWith('/ensureuser')).length, 2);
  const batch = sent.find((r) => r.url.endsWith('/_api/$batch'));
  assert.equal(batch.digestFrom, SITE);
  assert.match(batch.body, /addroleassignment\(principalid=31,/);
  assert.equal(result.batches, 1);
  assert.equal(result.unresolved.length, 1);
  const ops = result.lists[0].ops;
  assert.deepEqual(ops.map((o) => [o.outcome, o.label]), [
    ['success', 'Break inheritance (copy current permissions)'],
    ['success', 'Add Read: Ben'],
    ['failed', 'Add Read: Ghost'],
  ]);
  assert.match(ops[2].errorMessage, /could not be resolved \(HTTP 404: User not found\)/);
});

// ---------------------------------------------------------------- item-level permissions

const withSecurity = (read, write) => ({ unique: false, assignments: inherited, readSecurity: read, writeSecurity: write });

test('item-level only: no permission changes, also in replace mode; only changed fields are set', () => {
  const plan = planList(LIST, withSecurity(1, 1), [], { mode: 'replace', keepGroups: [] }, SITE_DATA, { read: 2, write: 1 });
  assert.deepEqual(plan.actions, []);
  assert.equal(plan.breaks, false);
  assert.deepEqual(plan.settings.map((s) => [s.field, s.change, s.from, s.to]), [
    ['ReadSecurity', true, 'Read all items', 'Read items that were created by the user'],
    ['WriteSecurity', false, 'Create and edit all items', 'Create and edit all items'],
  ]);
  assert.equal(countActions(plan).set, 1);
  assert.equal(hasChanges(plan), true);
  assert.equal(hasChanges(planList(LIST, withSecurity(2, 4), [], { mode: 'add' }, SITE_DATA, { read: 2, write: 4 })), false);
});

test('item-level setting is one MERGE on the list, before the permission operations, and parses', () => {
  const plan = planList(LIST, withSecurity(1, 1), [entry(members, READ)], { mode: 'add' }, SITE_DATA, { read: null, write: 4 });
  const ops = listOperations(SITE, plan);
  assert.deepEqual(ops.map((o) => o.method), ['MERGE', 'POST', 'POST']);
  assert.equal(ops[0].url, `${SITE}/_api/web/lists(guid'aaaa-1')`);
  assert.deepEqual(JSON.parse(ops[0].body), { __metadata: { type: 'SP.List' }, WriteSecurity: 4 });
  const parsed = parseBatch(buildPermissionBatch(ops));
  assert.deepEqual(parsed.errors, []);
  assert.equal(parsed.operations[0].method, 'MERGE');
  assert.equal(parsed.operations[0].headers['if-match'], '*');
  assert.equal(parsed.operations[0].body, ops[0].body);
});

test('libraries are recognised; item-level text', () => {
  assert.equal(isLibrary({ baseType: 1, template: 101 }), true);
  assert.equal(isLibrary({ template: 109 }), true);
  assert.equal(isLibrary({ baseType: 0, template: 100 }), false);
  assert.equal(itemLevelText({ read: 2, write: 4 }), 'Read: own items · Create and edit: none');
  assert.equal(itemLevelText({ read: null, write: 2 }), 'Create and edit: create, edit own');
});

test('replace mode keeps only the chosen default groups', () => {
  const plan = planList(LIST, { unique: true, assignments: inherited }, [entry(ben, READ)], { mode: 'replace', keepGroups: [3] }, SITE_DATA);
  assert.deepEqual(types(plan).filter((t) => !t.startsWith('add')), [
    'keep Demo Owners', 'remove Demo Members Edit', 'remove Demo Visitors Read', 'remove Project Group Edit', 'keep Anna',
  ]);
  assert.deepEqual(plan.warnings, []);
  // A non-default group in keepGroups is not kept.
  const other = planList(LIST, { unique: true, assignments: inherited }, [entry(ben, READ)], { mode: 'replace', keepGroups: [12] }, SITE_DATA);
  assert.ok(types(other).includes('remove Project Group Edit'));
});

test('loadRoles reads only the permission levels', async () => {
  const urls = [];
  const request = async (req) => {
    urls.push(req.url);
    return json(200, { value: [{ Id: 7, Name: 'Approve', RoleTypeKind: 0, Order: 64 }, { Id: 9, Name: 'Limited Access', Hidden: true }] });
  };
  const roles = await createPermissionClient({ request }).loadRoles(SITE);
  assert.deepEqual(urls, [`${SITE}/_api/web/roledefinitions?$select=Id,Name,Description,Hidden,RoleTypeKind,Order`]);
  assert.deepEqual(roles.map((r) => r.name), ['Approve']);
});
