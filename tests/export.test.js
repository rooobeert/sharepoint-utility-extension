import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  CSV_COLUMNS,
  csvCell,
  toCsv,
  csvFileName,
  isSystemList,
  pickerLists,
  isCustomField,
  columnRows,
  viewRows,
  contentTypeRows,
  permissionRows,
  listRow,
  siteRow,
} from '../extension/lib/export.js';
import { createExporter } from '../extension/lib/exporter.js';

const SITE = 'https://contoso.sharepoint.com/sites/demo';
const LIST = { Id: 'L1', Title: 'Change Requests', RootFolder: { ServerRelativeUrl: '/sites/demo/Lists/crRequests' } };

// ---------------------------------------------------------------- CSV

test('csvCell quotes separators, quotes and line breaks', () => {
  assert.equal(csvCell('plain', ';'), 'plain');
  assert.equal(csvCell('a;b', ';'), '"a;b"');
  assert.equal(csvCell('a,b', ';'), 'a,b');
  assert.equal(csvCell('a,b', ','), '"a,b"');
  assert.equal(csvCell('say "hi"', ';'), '"say ""hi"""');
  assert.equal(csvCell('<Field>\n</Field>', ';'), '"<Field>\n</Field>"');
  assert.equal(csvCell(null, ';'), '');
  assert.equal(csvCell(0, ';'), '0');
});

test('csvCell defuses values Excel would run as formulas', () => {
  assert.equal(csvCell('=HYPERLINK("http://x")', ';'), `"'=HYPERLINK(""http://x"")"`);
  assert.equal(csvCell('+1', ';'), "'+1");
  assert.equal(csvCell('-x', ';'), "'-x");
  assert.equal(csvCell('@SUM(A1)', ';'), "'@SUM(A1)");
  assert.equal(csvCell('0x0100ABC', ';'), '0x0100ABC');
  assert.equal(csvCell('<Field />', ';'), '<Field />');
});

test('toCsv: BOM, header, CRLF rows, semicolon by default, multi-line cells stay one record', () => {
  const csv = toCsv([{ Type: 'ListColumn', Name: 'crStatus', SchemaXml: '<Field\nType="Choice" />' }]);
  assert.ok(csv.startsWith('﻿Type;List;ListUrl;Name;Title;Id;Kind;Hidden;Required;Roles;Description;Settings;SchemaXml\r\n'));
  assert.ok(csv.endsWith('\r\n'));
  const record = csv.slice(1).split('\r\n')[1];
  assert.equal(record, 'ListColumn;;;crStatus;;;;;;;;;"<Field\nType=""Choice"" />"');
  assert.equal(toCsv([], ',').slice(1, 12), 'Type,List,L');
  assert.equal(CSV_COLUMNS.length, 13);
});

test('csvFileName is safe for Windows', () => {
  assert.equal(csvFileName('R&D: Hub/Projects?', new Date('2026-10-09T10:00:00Z')), 'R&D-HubProjects-architecture-2026-10-09.csv');
  assert.equal(csvFileName('', new Date('2026-10-09T10:00:00Z')), 'site-architecture-2026-10-09.csv');
});

// ---------------------------------------------------------------- picker

test('system lists: hidden, catalogs, IsSystemList', () => {
  assert.equal(isSystemList({ Hidden: true }), true);
  assert.equal(isSystemList({ IsCatalog: true }), true);
  assert.equal(isSystemList({ IsSystemList: true }), true);
  assert.equal(isSystemList({ RootFolder: { ServerRelativeUrl: '/sites/demo/_catalogs/masterpage' } }), true);
  assert.equal(isSystemList({ RootFolder: { ServerRelativeUrl: '/sites/demo/Lists/crRequests' } }), false);
});

test('pickerLists: compact, sorted by title, GUIDs normalised', () => {
  const lists = pickerLists({
    value: [
      { Id: '{BBB}', Title: 'b list', BaseTemplate: 100, ItemCount: 3, HasUniqueRoleAssignments: true, RootFolder: { ServerRelativeUrl: '/s/Lists/b' } },
      { Id: 'AAA', Title: 'A lib', BaseTemplate: 101, ItemCount: 0, Hidden: false, RootFolder: { ServerRelativeUrl: '/s/A' } },
    ],
  });
  assert.deepEqual(lists.map((l) => [l.id, l.title, l.kind, l.uniquePermissions]), [
    ['aaa', 'A lib', 'DocumentLibrary (101)', undefined],
    ['bbb', 'b list', 'GenericList (100)', true],
  ]);
});

// ---------------------------------------------------------------- rows

const FIELDS = {
  value: [
    { Id: 'f-title', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', FromBaseType: true, Required: true, SchemaXml: '<Field Name="Title" />' },
    { Id: 'f-created', InternalName: 'Created', Title: 'Created', TypeAsString: 'DateTime', FromBaseType: true, ReadOnlyField: true },
    { Id: 'f-hidden', InternalName: '_ComplianceFlags', Title: 'Label', TypeAsString: 'Text', Hidden: true },
    { Id: 'f-status', InternalName: 'crStatus', Title: 'Status', TypeAsString: 'Choice', Required: false, Choices: ['New', 'Done'], DefaultValue: 'New', Indexed: true, SchemaXml: '<Field Type="Choice" Name="crStatus" />' },
    { Id: '{F-LOOKUP}', InternalName: 'crSite', Title: 'Site', TypeAsString: 'Lookup', LookupList: '{AAAA-1}', LookupField: 'Title', SchemaXml: '<Field Type="Lookup" />' },
  ],
};

test('custom columns only by default; Title always included', () => {
  assert.deepEqual(columnRows(LIST, FIELDS).map((r) => r.Name), ['Title', 'crStatus', 'crSite']);
  assert.deepEqual(columnRows(LIST, FIELDS, { allColumns: true }).map((r) => r.Name), ['Title', 'Created', '_ComplianceFlags', 'crStatus', 'crSite']);
  assert.equal(isCustomField({ InternalName: 'x', Sealed: true }), false);
});

test('column rows: kind, flags, settings JSON with choices and readable lookup target', () => {
  const rows = columnRows(LIST, FIELDS, { listsById: { 'aaaa-1': { title: 'Production Sites', url: '/sites/demo/Lists/crSites' } } });
  const status = rows[1];
  assert.equal(status.Type, 'ListColumn');
  assert.equal(status.List, 'Change Requests');
  assert.equal(status.ListUrl, '/sites/demo/Lists/crRequests');
  assert.equal(status.Kind, 'Choice');
  assert.equal(status.Required, 'FALSE');
  assert.deepEqual(JSON.parse(status.Settings), { DefaultValue: 'New', Indexed: true, Choices: ['New', 'Done'] });
  assert.equal(status.SchemaXml, '<Field Type="Choice" Name="crStatus" />');
  const lookup = rows[2];
  assert.equal(lookup.Id, 'f-lookup');
  assert.deepEqual(JSON.parse(lookup.Settings), { LookupList: 'Production Sites', LookupListUrl: '/sites/demo/Lists/crSites', LookupField: 'Title' });
});

test('view rows: default view marked, fields in order, ListViewXml as SchemaXml', () => {
  const rows = viewRows(LIST, {
    value: [{ Id: 'V1', Title: 'All Items', ServerRelativeUrl: '/sites/demo/Lists/crRequests/AllItems.aspx', ViewType: 'HTML', DefaultView: true, RowLimit: 30, Paged: true, ViewQuery: '<OrderBy />', ViewFields: { Items: ['LinkTitle', 'crStatus'] }, ListViewXml: '<View />' }],
  });
  assert.equal(rows[0].Kind, 'HTML (default)');
  assert.deepEqual(JSON.parse(rows[0].Settings), { DefaultView: true, ViewFields: ['LinkTitle', 'crStatus'], RowLimit: 30, Paged: true, ViewQuery: '<OrderBy />' });
  assert.equal(rows[0].SchemaXml, '<View />');
});

test('content type rows: ID, group, field links', () => {
  const rows = contentTypeRows(LIST, { value: [{ Name: 'Item', StringId: '0x0100ABC', Group: 'List Content Types', FieldLinks: [{ Name: 'Title' }, { Name: 'crStatus' }], SchemaXml: '<ContentType />' }] });
  assert.deepEqual([rows[0].Name, rows[0].Title, rows[0].Kind], ['0x0100ABC', 'Item', 'List Content Types']);
  assert.deepEqual(JSON.parse(rows[0].Settings), { Fields: ['Title', 'crStatus'] });
});

test('permission rows: hidden roles (Limited Access) dropped, principal type named', () => {
  const rows = permissionRows(
    {
      value: [
        { PrincipalId: 3, Member: { Id: 3, LoginName: 'Demo Owners', Title: 'Demo Owners', PrincipalType: 8 }, RoleDefinitionBindings: [{ Name: 'Full Control' }] },
        { PrincipalId: 9, Member: { Id: 9, LoginName: 'i:0#.f|membership|anna@contoso.com', Title: 'Anna', PrincipalType: 1, Email: 'anna@contoso.com' }, RoleDefinitionBindings: [{ Name: 'Edit' }, { Name: 'Limited Access', Hidden: true }] },
        { PrincipalId: 11, Member: { Id: 11, LoginName: 'x', Title: 'Only limited', PrincipalType: 1 }, RoleDefinitionBindings: [{ Name: 'Limited Access', Hidden: true }] },
      ],
    },
    { type: 'ListPermission', list: LIST, inherited: false }
  );
  assert.deepEqual(rows.map((r) => [r.Title, r.Kind, r.Roles]), [
    ['Demo Owners', 'SharePointGroup', 'Full Control'],
    ['Anna', 'User', 'Edit'],
  ]);
  assert.deepEqual(JSON.parse(rows[1].Settings), { Inherited: false, Email: 'anna@contoso.com' });
  assert.equal(rows[0].List, 'Change Requests');
});

test('list and site rows', () => {
  const l = listRow({ ...LIST, BaseTemplate: 101, Hidden: false, ItemCount: 4, EnableVersioning: true, MajorVersionLimit: 500, EnableMinorVersions: false, HasUniqueRoleAssignments: true, DefaultView: { ServerRelativeUrl: '/x/AllItems.aspx' } });
  assert.equal(l.Kind, 'DocumentLibrary (101)');
  assert.equal(l.Hidden, 'FALSE');
  assert.deepEqual(JSON.parse(l.Settings), { ItemCount: 4, EnableVersioning: true, MajorVersionLimit: 500, HasUniqueRoleAssignments: true, DefaultViewUrl: '/x/AllItems.aspx' });
  const s = siteRow({ Title: 'Demo', Url: SITE, Id: '{W}', WebTemplate: 'GROUP', Configuration: 0, Language: 1031, ServerRelativeUrl: '/sites/demo' }, SITE);
  assert.deepEqual([s.Type, s.Name, s.Kind, s.Id], ['Site', SITE, 'GROUP#0', 'w']);
});

// ---------------------------------------------------------------- exporter with a fake SharePoint

function fakeSharePoint({ denyListPermissions = false, rejectIsSystemList = false, throttleOnce = false } = {}) {
  const calls = [];
  let throttled = false;
  const ok = (body) => ({ status: 200, headers: {}, text: JSON.stringify(body) });
  const request = async (req) => {
    const path = decodeURIComponent(req.url.slice(SITE.length));
    calls.push(req.method + ' ' + path);
    assert.equal(req.method, 'GET');
    if (throttleOnce && !throttled) {
      throttled = true;
      return { status: 429, headers: { 'retry-after': '2' }, text: '' };
    }
    if (path.startsWith('/_api/web?')) return ok({ Title: 'Demo', Url: SITE, Id: 'W', ServerRelativeUrl: '/sites/demo', WebTemplate: 'STS', Configuration: 3 });
    if (path.startsWith('/_api/site?')) return ok({ Url: SITE });
    if (path.startsWith('/_api/web/lists?')) {
      if (rejectIsSystemList && path.includes('IsSystemList')) return { status: 400, headers: {}, text: '{"odata.error":{"message":{"value":"The property IsSystemList does not exist."}}}' };
      return ok({ value: [{ ...LIST, Id: 'AAAA-1', BaseTemplate: 100 }, { Id: 'BBBB-2', Title: 'Style Library', BaseTemplate: 101, IsCatalog: false, Hidden: false, IsSystemList: true, RootFolder: { ServerRelativeUrl: '/sites/demo/Style Library' } }] });
    }
    if (path === '/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings') {
      return ok({ value: [{ Member: { Id: 3, LoginName: 'Demo Owners', Title: 'Demo Owners', PrincipalType: 8 }, RoleDefinitionBindings: [{ Name: 'Full Control' }] }] });
    }
    const list = /^\/_api\/web\/lists\(guid'aaaa-1'\)(.*)$/.exec(path);
    if (list) {
      const rest = list[1];
      if (rest.startsWith('?')) return ok({ ...LIST, Id: 'AAAA-1', BaseTemplate: 100, HasUniqueRoleAssignments: false });
      if (rest.startsWith('/roleassignments')) {
        return denyListPermissions ? { status: 403, headers: {}, text: '{"odata.error":{"message":{"value":"Access denied."}}}' } : ok({ value: [] });
      }
      if (rest === '/fields') return ok(FIELDS);
      if (rest.startsWith('/views')) return ok({ value: [{ Id: 'V1', Title: 'All Items', ViewType: 'HTML', DefaultView: true, ViewFields: { Items: ['LinkTitle'] } }] });
      if (rest.startsWith('/contenttypes')) return ok({ value: [{ Name: 'Item', StringId: '0x01', FieldLinks: [] }] });
    }
    return { status: 404, headers: {}, text: '{}' };
  };
  return { request, calls };
}

test('exporter: loads the picker, exports selected lists in type order, all GET', async () => {
  const sp = fakeSharePoint();
  const events = [];
  const exporter = createExporter({ request: sp.request, sleep: async () => {}, onEvent: (e) => events.push(e) });
  const loaded = await exporter.loadLists(SITE);
  assert.deepEqual(loaded.lists.map((l) => [l.title, l.system]), [['Change Requests', false], ['Style Library', true]]);

  const { rows, problems } = await exporter.exportSite(SITE, loaded, ['aaaa-1']);
  assert.deepEqual(problems, []);
  assert.deepEqual(rows.map((r) => r.Type), ['Site', 'SitePermission', 'List', 'ListColumn', 'ListColumn', 'ListColumn', 'ListView', 'ListContentType']);
  assert.equal(JSON.parse(rows.find((r) => r.Name === 'crSite').Settings).LookupList, 'Change Requests'); // lookup to itself, resolved by GUID
  assert.ok(!sp.calls.some((c) => c.includes("guid'bbbb-2'")));
  assert.ok(events.some((e) => e.type === 'list' && e.title === 'Change Requests'));
});

test('exporter: unreadable permissions become an error row, the rest is exported', async () => {
  const sp = fakeSharePoint({ denyListPermissions: true });
  const exporter = createExporter({ request: sp.request, sleep: async () => {} });
  const loaded = await exporter.loadLists(SITE);
  const { rows, problems } = await exporter.exportSite(SITE, loaded, ['aaaa-1']);
  assert.deepEqual(problems, ['Change Requests: Permissions could not be read (HTTP 403: Access denied.).']);
  const err = rows.find((r) => r.Type === 'ListPermission');
  assert.equal(err.Title, '(Permissions could not be read)');
  assert.ok(rows.some((r) => r.Type === 'ListView'));
});

test('exporter: retries on 429 and falls back when IsSystemList is unknown', async () => {
  const sp = fakeSharePoint({ rejectIsSystemList: true, throttleOnce: true });
  const waits = [];
  const exporter = createExporter({ request: sp.request, sleep: async (ms) => waits.push(ms) });
  const loaded = await exporter.loadLists(SITE);
  assert.deepEqual(waits, [2000]);
  assert.equal(loaded.lists.length, 2);
  assert.ok(sp.calls.some((c) => c.startsWith('GET /_api/web/lists?') && !c.includes('IsSystemList')));
});
