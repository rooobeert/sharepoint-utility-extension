import { test } from 'node:test';
import assert from 'node:assert/strict';
import { forTarget, settingsStage, columnStage, columnTransferable, settingsTransferable } from '../extension/lib/transfer.js';

const SITE = 'https://t.sharepoint.com/sites/prod';
const ctx = { web: { ServerRelativeUrl: '/sites/dev' }, allLists: [], targetSite: SITE };

test('forTarget writes the target URLs into the placeholders', () => {
  assert.equal(forTarget('{{WebUrl}}/_api/x {{WebServerRelativeUrl}}/Lists/a {{ListId:Lists/a}}', SITE),
    `${SITE}/_api/x /sites/prod/Lists/a {{ListId:Lists/a}}`);
});

test('settingsStage merges the source values of the changed settings', () => {
  const list = { title: 'Req', rel: 'Lists/Req', item: { list: { Title: 'Req', EnableVersioning: true, MajorVersionLimit: 50 } },
    diffs: [{ prop: 'Versioning', source: 'Yes', target: 'No' }, { prop: 'Template', blocked: true }] };
  assert.ok(settingsTransferable(list));
  const st = settingsStage(list, SITE);
  assert.match(st.body, /^MERGE https:\/\/t\.sharepoint\.com\/sites\/prod\/_api\/web\/getList\('\/sites\/prod\/Lists\/Req'\)/m);
  assert.match(st.body, /"EnableVersioning":true/);
  assert.doesNotMatch(st.body, /BaseTemplate|Template/);
});

test('columnStage: changed column merges typed properties, blocked ones are left out', () => {
  const list = { title: 'Req', rel: 'Lists/Req' };
  const col = { name: 'Notes', title: 'Notes', type: 'Note', status: 'changed', field: { TypeAsString: 'Note', Title: 'Notes', NumberOfLines: 6, Required: true },
    diffs: [{ prop: 'Lines', source: '6', target: '3' }, { prop: 'Required', source: 'Yes', target: 'No' }, { prop: 'Type', blocked: true }] };
  const st = columnStage(list, col, ctx);
  assert.match(st.body, /fields\/getByInternalNameOrTitle\('Notes'\)/);
  assert.match(st.body, /SP\.FieldMultiLineText/);
  assert.match(st.body, /"NumberOfLines":6/);
  assert.match(st.body, /"Required":true/);
});

test('columnStage: missing column is created from its cleaned schema', () => {
  const col = { name: 'Foo', title: 'Foo', type: 'Text', status: 'missing', field: { SchemaXml: '<Field Name="Foo" SourceID="x" Type="Text"/>' } };
  const st = columnStage({ title: 'Req', rel: 'Lists/Req' }, col, ctx);
  assert.match(st.body, /POST .*fields\/CreateFieldAsXml/);
  assert.doesNotMatch(st.body, /SourceID/);
});

test('nothing to send for managed metadata or blocked-only changes', () => {
  assert.ok(!columnTransferable({ status: 'missing', type: 'TaxonomyFieldType', field: { SchemaXml: '<Field/>' } }));
  assert.ok(!columnTransferable({ status: 'changed', field: {}, diffs: [{ prop: 'Type', blocked: true }] }));
});
