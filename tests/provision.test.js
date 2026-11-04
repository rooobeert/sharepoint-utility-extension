import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  buildProvisioning,
  cleanFieldXml,
  webRelative,
  splitStageFile,
  batchFileName,
  STAGE_MARKER,
} from '../extension/lib/provision.js';
import { parseBatch } from '../extension/lib/batch.js';
import { findPlaceholders, findUrlPlaceholders } from '../extension/lib/placeholders.js';
import { createRunner } from '../extension/lib/runner.js';

const SRU = '/sites/src';
const WEB = { Title: 'Source', Url: 'https://contoso.sharepoint.com/sites/src', ServerRelativeUrl: SRU };
const ALL = [
  { id: 'aaaa-1', title: 'Change Requests', url: `${SRU}/Lists/crRequests` },
  { id: 'bbbb-2', title: 'Departments', url: `${SRU}/Lists/crDepartments` },
  { id: 'cccc-3', title: 'Documents', url: `${SRU}/Shared Documents` },
  { id: 'dddd-4', title: 'CR Attachments', url: `${SRU}/crAttachments` },
];
const val = (value) => ({ value });

const requests = {
  entry: { id: 'aaaa-1', title: 'Change Requests' },
  list: {
    Title: 'Change Requests', Description: 'All changes', BaseTemplate: 100, BaseType: 0,
    RootFolder: { ServerRelativeUrl: `${SRU}/Lists/crRequests` },
    EnableVersioning: true, MajorVersionLimit: 50, EnableAttachments: false, EnableFolderCreation: false, ContentTypesEnabled: false,
  },
  fields: val([
    { InternalName: 'Title', Title: 'Change-ID', Required: true, FromBaseType: true },
    { InternalName: 'Created', Title: 'Created', FromBaseType: true, ReadOnlyField: true, SchemaXml: '<Field Name="Created" />' },
    { InternalName: 'crStatus', Title: 'Status', TypeAsString: 'Choice', SchemaXml: '<Field ID="{11111111-2222-4333-8444-555555555555}" Type="Choice" Name="crStatus" DisplayName="Status" SourceID="{0b1c}" StaticName="crStatus" ColName="nvarchar4" RowOrdinal="0" Version="3"><CHOICES><CHOICE>Neu</CHOICE></CHOICES></Field>' },
    { InternalName: 'crDept', Title: 'Department', TypeAsString: 'Lookup', SchemaXml: '<Field ID="{22222222-2222-4333-8444-555555555555}" Type="Lookup" Name="crDept" List="{BBBB-2}" WebId="{9999}" ShowField="Title" SourceID="{0b1c}" />' },
    { InternalName: 'crTags', Title: 'Tags', TypeAsString: 'TaxonomyFieldTypeMulti', SchemaXml: '<Field Type="TaxonomyFieldTypeMulti" />' },
  ]),
  views: val([
    { Title: 'Alle Elemente', DefaultView: true, ViewType: 'HTML', ServerRelativeUrl: `${SRU}/Lists/crRequests/AllItems.aspx`, ViewFields: { Items: ['LinkTitle', 'crStatus'] }, ViewQuery: '<OrderBy><FieldRef Name="ID" /></OrderBy>', RowLimit: 30, Paged: true },
    { Title: 'Offene Änderungen', ViewType: 'HTML', ServerRelativeUrl: `${SRU}/Lists/crRequests/Open.aspx`, ViewFields: { Items: ['LinkTitle'] }, ViewQuery: '<Where />', RowLimit: 100, Paged: true, Scope: 1 },
    { Title: 'Kalender', ViewType: 'CALENDAR', ServerRelativeUrl: `${SRU}/Lists/crRequests/calendar.aspx`, ViewFields: { Items: [] } },
    { Title: 'Hidden', Hidden: true, ViewType: 'HTML', ServerRelativeUrl: `${SRU}/Lists/crRequests/h.aspx`, ViewFields: { Items: [] } },
  ]),
  errors: {},
};
const departments = {
  entry: { id: 'bbbb-2', title: 'Departments' },
  list: { Title: 'Departments', BaseTemplate: 100, BaseType: 0, RootFolder: { ServerRelativeUrl: `${SRU}/Lists/crDepartments` } },
  fields: val([{ InternalName: 'Title', Title: 'Title', Required: true, FromBaseType: true }]),
  views: val([]),
  errors: {},
};
const attachments = {
  entry: { id: 'dddd-4', title: 'CR Attachments' },
  list: { Title: 'CR Attachments', BaseTemplate: 101, BaseType: 1, RootFolder: { ServerRelativeUrl: `${SRU}/crAttachments` }, EnableVersioning: true, MajorVersionLimit: 500, EnableMinorVersions: false, ForceCheckout: false, EnableAttachments: false },
  fields: val([{ InternalName: 'Title', Title: 'Title', Required: false, FromBaseType: true }, { InternalName: 'crDoc', Title: 'Doc', TypeAsString: 'Text', SchemaXml: '<Field Type="Text" Name="crDoc" />' }]),
  views: val([]),
  errors: {},
};
const documents = {
  entry: { id: 'cccc-3', title: 'Documents' },
  list: { Title: 'Documents', BaseTemplate: 101, BaseType: 1, RootFolder: { ServerRelativeUrl: `${SRU}/Shared Documents` } },
  fields: val([{ InternalName: 'crReq', Title: 'Request', TypeAsString: 'Lookup', SchemaXml: '<Field Type="Lookup" Name="crReq" List="{aaaa-1}" ShowField="Title" />' }]),
  views: val([]),
  errors: {},
};

const build = (opts = {}) => buildProvisioning({ web: WEB, allLists: ALL, lists: [requests, departments, attachments, documents], version: '1.2.0', date: new Date('2026-10-09T10:00:00Z'), ...opts });

test('webRelative', () => {
  assert.equal(webRelative('/sites/src/Lists/crRequests', '/sites/src'), 'Lists/crRequests');
  assert.equal(webRelative('/sites/SRC/Shared Documents', '/sites/src/'), 'Shared Documents');
  assert.equal(webRelative('/Lists/x', '/'), 'Lists/x');
  assert.equal(webRelative('/Lists/x', ''), 'Lists/x');
});

test('cleanFieldXml removes source-site attributes, keeps ID, turns lookups into placeholders', () => {
  const lists = { 'bbbb-2': { title: 'Departments', url: `${SRU}/Lists/crDepartments` } };
  const { xml, notes } = cleanFieldXml('<Field ID="{2222}" Type="Lookup" Name="crDept" List="{BBBB-2}" WebId="{9}" SourceID="{0}" ColName="int1" RowOrdinal="0" Version="2" ShowField="Title" />', lists, SRU);
  assert.equal(xml, '<Field ID="{2222}" Type="Lookup" Name="crDept" List="{{{ListId:Lists/crDepartments}}}" ShowField="Title" />');
  assert.deepEqual(notes, []);
  assert.equal(cleanFieldXml('<Field List="Self" />', lists, SRU).xml, '<Field List="Self" />');
  assert.match(cleanFieldXml('<Field List="{eeee}" />', lists, SRU).notes[0], /not a list of this site/);
});

test('stages, in order, with the expected operations', () => {
  const { stages } = build();
  assert.deepEqual(stages.map((s) => [s.name, s.ops.length]), [
    ['01 Lists and libraries', 3], // Documents exists on every site
    ['02 List settings', 2], // Departments has no settings to set
    ['03 Columns', 5], // Title MERGE, crStatus, crDept on requests; crDoc; crReq on Documents
    ['04 Views', 6], // default view: MERGE + removeall + 2 fields; Open view add + Scope MERGE
    ['05 Display names', 4], // 3 lists renamed, 1 view renamed
  ]);
  const create = stages[0].ops.map((o) => JSON.parse(o.body));
  assert.deepEqual(create.map((b) => [b.Title, b.BaseTemplate]), [['crRequests', 100], ['crDepartments', 100], ['crAttachments', 101]]);
});

test('list settings match the list type', () => {
  const [req, lib] = build().stages[1].ops.map((o) => JSON.parse(o.body));
  assert.deepEqual(req, { __metadata: { type: 'SP.List' }, EnableVersioning: true, MajorVersionLimit: 50, ContentTypesEnabled: false, EnableFolderCreation: false, EnableAttachments: false });
  assert.deepEqual(lib, { __metadata: { type: 'SP.List' }, EnableVersioning: true, MajorVersionLimit: 500, EnableMinorVersions: false, ForceCheckout: false });
});

test('columns: cleaned SchemaXml with field IDs, Title change, lookup placeholder, taxonomy skipped', () => {
  const { stages, notes } = build();
  const cols = stages[2].ops;
  assert.match(cols[0].url, /fields\/getByInternalNameOrTitle\('Title'\)$/);
  assert.deepEqual(JSON.parse(cols[0].body), { __metadata: { type: 'SP.Field' }, Title: 'Change-ID', Required: true });
  const status = JSON.parse(cols[1].body).parameters;
  assert.equal(status.SchemaXml, '<Field ID="{11111111-2222-4333-8444-555555555555}" Type="Choice" Name="crStatus" DisplayName="Status" StaticName="crStatus"><CHOICES><CHOICE>Neu</CHOICE></CHOICES></Field>');
  assert.equal(status.Options, 24);
  assert.match(JSON.parse(cols[2].body).parameters.SchemaXml, /List="\{\{\{ListId:Lists\/crDepartments\}\}\}"/);
  assert.match(JSON.parse(cols[4].body).parameters.SchemaXml, /List="\{\{\{ListId:Lists\/crRequests\}\}\}"/);
  assert.ok(notes.some((n) => /managed metadata column "Tags"/.test(n)));
  assert.ok(notes.some((n) => /"Documents" \(Shared Documents\) exists on every site/.test(n)));
  assert.ok(notes.some((n) => /view "Kalender" \(CALENDAR\) skipped/.test(n)));
});

test('views: default view changed in place, others created under their URL name and renamed later', () => {
  const { stages } = build();
  const v = stages[3].ops;
  assert.equal(v[0].method, 'MERGE');
  assert.match(v[0].url, /\/defaultView$/);
  assert.match(v[1].url, /\/defaultView\/viewfields\/removeallviewfields$/);
  assert.deepEqual(v.slice(2, 4).map((o) => o.url.split('/').pop()), ["addviewfield('LinkTitle')", "addviewfield('crStatus')"]);
  const add = JSON.parse(v[4].body).parameters;
  assert.equal(add.Title, 'Open');
  assert.deepEqual(add.ViewFields.results, ['LinkTitle']);
  assert.equal(add.Query, '<Where />');
  assert.match(v[5].url, /views\/getByTitle\('Open'\)$/);
  const rename = stages[4].ops.find((o) => /getByTitle\('Open'\)/.test(o.url));
  assert.equal(JSON.parse(rename.body).Title, 'Offene Änderungen');
});

test('bodyless operations carry no Content-Type (addviewfield would fail on empty JSON)', () => {
  const { text } = build();
  const parts = text.split(/^--changeset_provisioning\n/m).slice(1);
  const addField = parts.find((p) => p.includes("addviewfield('LinkTitle')"));
  assert.doesNotMatch(addField, /application\/json;odata=verbose/);
  const merge = parts.find((p) => /^MERGE .*\/defaultView HTTP/m.test(p));
  assert.match(merge, /Content-Type: application\/json;odata=verbose/);
});

test('URLs: placeholders only, spaces encoded, quotes doubled', () => {
  const { stages } = build();
  const urls = stages.flatMap((s) => s.ops.map((o) => o.url));
  assert.ok(urls.every((u) => u.startsWith('{{WebUrl}}/_api/web/')));
  assert.ok(urls.some((u) => u.includes("getList('{{WebServerRelativeUrl}}/Shared%20Documents')")));
  assert.ok(urls.every((u) => !/\s/.test(u)));
  const quoted = buildProvisioning({
    web: WEB, allLists: [], lists: [{ entry: { title: "Bob's" }, list: { Title: "Bob's", BaseTemplate: 100, RootFolder: { ServerRelativeUrl: `${SRU}/Lists/Bob's #1` } }, errors: {} }],
  });
  assert.match(quoted.stages.at(-1).ops[0].url, /getList\('\{\{WebServerRelativeUrl\}\}\/Lists\/Bob''s%20%231'\)$/);
});

test('file: header comments, stage markers, each stage a valid batch; splitStageFile reads it back', () => {
  const { text, stages } = build();
  assert.ok(text.startsWith('# SharePoint Utility: provisioning batches\n# Source: Source (https://contoso.sharepoint.com/sites/src), exported 2026-10-09 with v1.2.0\n'));
  assert.equal(text.split('\n').filter((l) => l.startsWith(STAGE_MARKER)).length, stages.length);
  const back = splitStageFile(text);
  assert.deepEqual(back.map((s) => s.name), stages.map((s) => s.name));
  for (const [i, s] of back.entries()) {
    const parsed = parseBatch(s.body);
    assert.deepEqual(parsed.errors, [], s.name);
    assert.equal(parsed.operations.length, stages[i].ops.length, s.name);
    assert.ok(findUrlPlaceholders(s.body).includes('WebUrl'));
  }
  assert.deepEqual(findPlaceholders(back[2].body).map((p) => p.ref), ['Lists/crDepartments', 'Lists/crRequests']);
});

test('splitStageFile: a file without markers is one stage; CRLF and BOM tolerated', () => {
  assert.deepEqual(splitStageFile('--batch_a\n--batch_a--', 'Batch1'), [{ name: 'Batch1', body: '--batch_a\n--batch_a--' }]);
  const two = splitStageFile('\uFEFF# c\r\n### STAGE: A\r\n--b\r\n### STAGE:\r\n--c\r\n');
  assert.deepEqual(two, [{ name: 'A', body: '--b\n' }, { name: 'Stage 2', body: '--c\n' }]);
});

test('large stages are split into parts of at most N operations', () => {
  const { stages } = build({ max: 2 });
  assert.deepEqual(stages.filter((s) => s.name.startsWith('03')).map((s) => [s.name, s.ops.length]), [
    ['03 Columns (1/3)', 2], ['03 Columns (2/3)', 2], ['03 Columns (3/3)', 1],
  ]);
});

test('batchFileName', () => {
  assert.equal(batchFileName('R&D: Hub', new Date('2026-10-09T10:00:00Z')), 'R&D-Hub-batches-2026-10-09.txt');
});

// ---------------------------------------------------------------- round trip through the Batch Runner

test('the generated file runs in the Batch Runner on an empty site: lists, lookups, views, names', async () => {
  const TARGET = 'https://contoso.sharepoint.com/sites/target';
  const TSRU = '/sites/target';
  const lists = [{ id: 'docs-guid', title: 'Documents', url: `${TSRU}/Shared Documents` }];
  const sent = [];
  let n = 0;
  const json = (body) => ({ status: 200, statusText: 'OK', headers: {}, text: JSON.stringify(body) });
  const request = async (req) => {
    if (req.method === 'GET') {
      const path = decodeURIComponent(req.url.slice(TARGET.length));
      if (path.startsWith('/_api/web/lists?')) return json({ value: lists.map((l) => ({ Id: l.id, Title: l.title, RootFolder: { ServerRelativeUrl: l.url } })) });
      if (path.startsWith('/_api/web?')) return json({ Id: 'web', Title: 'Target', Url: TARGET, ServerRelativeUrl: TSRU });
      if (path.startsWith('/_api/site?')) return json({ Url: TARGET });
      return { status: 404, headers: {}, text: '{}' };
    }
    const ops = parseBatch(req.body).operations;
    sent.push(...ops);
    let out = '';
    for (const o of ops) {
      if (/\/_api\/web\/lists$/.test(o.url)) {
        const b = JSON.parse(o.body);
        lists.push({ id: `new-${++n}`, title: b.Title, url: `${TSRU}${b.BaseTemplate === 101 ? '/' : '/Lists/'}${b.Title}` });
      }
      out += '--r\r\nContent-Type: application/http\r\n\r\nHTTP/1.1 201 Created\r\n\r\n{}\r\n';
    }
    return { status: 200, statusText: 'OK', headers: { 'content-type': 'multipart/mixed; boundary=r' }, text: out + '--r--\r\n' };
  };
  const runner = createRunner({ request, sleep: async () => {} });
  const run = await runner.runStages(TARGET, splitStageFile(build().text));

  assert.deepEqual(run.stages.map((s) => [s.status, s.summary.failed]), run.stages.map(() => ['completed', 0]));
  assert.ok(sent.every((o) => o.url.startsWith(`${TARGET}/_api/web/`)));
  assert.ok(sent.some((o) => o.url.includes(`getList('${TSRU}/Lists/crRequests')/fields/CreateFieldAsXml`)));
  const dept = sent.find((o) => o.body.includes('Name=\\"crDept\\"'));
  assert.match(dept.body, /List=\\"\{new-2\}\\"/); // crDepartments was the 2nd list created
  const req = sent.find((o) => o.body.includes('Name=\\"crReq\\"'));
  assert.match(req.body, /List=\\"\{new-1\}\\"/);
});
