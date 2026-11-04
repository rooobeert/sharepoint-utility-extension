import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  compareSites,
  listProps,
  columnProps,
  propDiffs,
  listStatus,
  countStatuses,
  columnSummary,
  filterView,
  diffRows,
  diffCsv,
  diffFileName,
  DIFF_CSV_COLUMNS,
} from '../extension/lib/diff.js';

const val = (value) => ({ value });

// Two sites in different site collections: the same lists have different GUIDs and server-relative URLs.
function side(sru, title, lists) {
  return {
    web: { Title: title, ServerRelativeUrl: sru },
    allLists: lists.map((l) => ({ id: l.id, title: l.title, url: `${sru}/${l.rel}` })),
    lists: lists.map((l) => ({
      entry: { id: l.id, title: l.title, url: `${sru}/${l.rel}` },
      list: l.error === 'List' ? null : {
        Title: l.title, BaseTemplate: l.template ?? 100, BaseType: l.template === 101 ? 1 : 0,
        RootFolder: { ServerRelativeUrl: `${sru}/${l.rel}` },
        EnableVersioning: true, MajorVersionLimit: l.versions ?? 50, EnableAttachments: true, EnableFolderCreation: false,
        ContentTypesEnabled: false, EnableModeration: false, NoCrawl: false,
      },
      fields: l.error ? null : val(l.fields || []),
      errors: l.error ? { [l.error === 'List' ? 'List' : 'ListColumn']: 'HTTP 403: Access denied' } : {},
    })),
  };
}

const title = { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Required: true, FromBaseType: true, Id: '{fa564e0f-0c70-4ab9-b863-0177e6ddd247}' };
const created = { InternalName: 'Created', Title: 'Created', TypeAsString: 'DateTime', FromBaseType: true, ReadOnlyField: true, Id: '8c06beca' };
const field = (name, props = {}) => ({ InternalName: name, Title: name.replace(/^cr/, ''), TypeAsString: 'Text', MaxLength: 255, Required: false, Id: `{id-${name}}`, ...props });

function sites() {
  const source = side('/sites/dev', 'R&D Hub DEV', [
    {
      id: 's-req', title: 'Requests', rel: 'Lists/crRequests', versions: 500,
      fields: [
        title, created,
        field('crPriority', { TypeAsString: 'Choice', Required: true, Choices: ['Low', 'High', 'Urgent'], MaxLength: undefined }),
        field('crDueDate', { TypeAsString: 'DateTime', MaxLength: undefined }),
        field('crCustomer', { TypeAsString: 'Lookup', LookupList: '{S-CUST}', LookupField: 'Title', AllowMultipleValues: false, MaxLength: undefined }),
        field('crBudget', { TypeAsString: 'Currency', MaxLength: undefined }),
        field('crNotes'),
      ],
    },
    { id: 's-cust', title: 'Customers', rel: 'Lists/crCustomers', fields: [title, field('crCity')] },
    { id: 's-new', title: 'Change Requests', rel: 'Lists/crChangeRequests', fields: [title, field('crA'), field('crB'), created] },
    { id: 's-appr', title: 'Approvals', rel: 'Lists/crApprovals', error: 'ListColumn' },
    { id: 's-dup', title: 'Contacts', rel: 'Lists/crContactsNew', fields: [title] },
    { id: 's-docs', title: 'Documents', rel: 'Shared Documents', template: 101, fields: [title] },
  ]);
  const target = side('/sites/prod', 'R&D Hub', [
    {
      id: 't-req', title: 'Requests', rel: 'Lists/crRequests', versions: 100,
      fields: [
        title, created,
        field('crPriority', { TypeAsString: 'Choice', Required: false, Choices: ['Low', 'High'], MaxLength: undefined }),
        field('crDueDate'),
        field('crNotes'),
        field('crLegacyId'),
      ],
    },
    { id: 't-old', title: 'Old Tasks', rel: 'Lists/OldTasks', fields: [title, field('crX')] },
    { id: 't-appr', title: 'Approvals', rel: 'Lists/crApprovals', fields: [title] },
    { id: 't-dup', title: 'Contacts', rel: 'Lists/crContacts', fields: [title] },
  ]);
  return { source, target };
}

const byRel = (lists) => Object.fromEntries(lists.map((l) => [l.rel, l]));

test('lists are matched by web-relative URL across site collections', () => {
  const { source, target } = sites();
  const lists = byRel(compareSites(source, target));
  assert.equal(lists['Lists/crRequests'].status, 'changed');
  assert.equal(lists['Lists/crChangeRequests'].status, 'missing');
  assert.equal(lists['Lists/crCustomers'].status, 'missing');
  assert.equal(lists['Lists/OldTasks'].status, 'only');
  assert.equal(lists['Lists/crChangeRequests'].columnCount, 3, 'Title and the two custom columns, not Created');
});

test('a list whose columns could not be read on one side is not compared', () => {
  const { source, target } = sites();
  const appr = byRel(compareSites(source, target))['Lists/crApprovals'];
  assert.equal(appr.status, 'error');
  assert.match(appr.error, /columns could not be read in source.*403/);
  assert.deepEqual(appr.columns, []);
});

test('a list that could not be read at all is not compared', () => {
  const source = side('/sites/dev', 'Dev', [{ id: 'a', title: 'A', rel: 'Lists/A', error: 'List' }]);
  const target = side('/sites/prod', 'Prod', [{ id: 'b', title: 'A', rel: 'Lists/A', fields: [title] }]);
  const [a] = compareSites(source, target);
  assert.equal(a.status, 'error');
  assert.match(a.error, /the list could not be read in source/);
});

test('default libraries are compared only when both sides have them', () => {
  const { source, target } = sites();
  assert.equal(byRel(compareSites(source, target))['Shared Documents'], undefined);
  const both = sites();
  both.target.lists.push(...side('/sites/prod', 'x', [{ id: 't-docs', title: 'Dokumente', rel: 'Shared Documents', template: 101, fields: [title] }]).lists);
  const docs = byRel(compareSites(both.source, both.target))['Shared Documents'];
  assert.equal(docs.status, 'changed', 'title differs: Documents vs Dokumente');
  assert.equal(docs.title, 'Dokumente', 'a matched list shows the target title');
});

test('lists with the same title under different URLs get a hint on both sides', () => {
  const { source, target } = sites();
  const lists = byRel(compareSites(source, target));
  assert.match(lists['Lists/crContactsNew'].note, /same title at Lists\/crContacts$/);
  assert.match(lists['Lists/crContacts'].note, /same title at Lists\/crContactsNew$/);
});

test('list settings differences, template change is blocked', () => {
  const { source, target } = sites();
  const req = byRel(compareSites(source, target))['Lists/crRequests'];
  assert.deepEqual(req.diffs, [{ prop: 'Versioning limit', source: '500', target: '100' }]);

  const a = listProps({ BaseTemplate: 100, BaseType: 0, EnableVersioning: true, MajorVersionLimit: 50 });
  const b = listProps({ BaseTemplate: 107, BaseType: 0, EnableVersioning: true, MajorVersionLimit: 50 });
  const [tpl] = propDiffs(a, b, new Set(['Template']));
  assert.equal(tpl.prop, 'Template');
  assert.equal(tpl.blocked, true);
});

test('version limits only count while versioning is on; library and list settings differ', () => {
  assert.equal(listProps({ EnableVersioning: false, MajorVersionLimit: 50 })['Versioning limit'], '');
  const lib = listProps({ BaseType: 1, BaseTemplate: 101, ForceCheckout: true });
  assert.equal(lib['Require check-out'], 'Yes');
  assert.equal('Attachments' in lib, false);
  assert.equal('Require check-out' in listProps({ BaseType: 0, BaseTemplate: 100 }), false);
});

test('columns: missing, only in target, changed, same; sorted by status', () => {
  const { source, target } = sites();
  const req = byRel(compareSites(source, target))['Lists/crRequests'];
  const cols = Object.fromEntries(req.columns.map((c) => [c.name, c]));
  assert.equal(cols.crBudget.status, 'missing');
  assert.equal(cols.crCustomer.status, 'missing');
  assert.equal(cols.crLegacyId.status, 'only');
  assert.equal(cols.crNotes.status, 'same');
  assert.equal(cols.Created.custom, false);
  assert.equal(cols.Title.custom, true);

  const prio = cols.crPriority;
  assert.equal(prio.status, 'changed');
  const d = Object.fromEntries(prio.diffs.map((x) => [x.prop, x]));
  assert.deepEqual([d.Required.source, d.Required.target], ['Yes', 'No']);
  assert.deepEqual(d.Choices.added, ['Urgent'], 'the target would gain "Urgent"');
  assert.deepEqual(d.Choices.removed, []);

  const due = cols.crDueDate;
  const type = due.diffs.find((x) => x.prop === 'Type');
  assert.deepEqual([type.source, type.target, type.blocked], ['DateTime', 'Text', true]);
  assert.deepEqual(due.diffs.map((x) => x.prop), ['Type'], 'max length of the old Text type is not reported');

  assert.deepEqual(req.columns.map((c) => c.status), [...req.columns.map((c) => c.status)].sort(
    (a, b) => ['missing', 'changed', 'only', 'same'].indexOf(a) - ['missing', 'changed', 'only', 'same'].indexOf(b)));
});

test('a missing lookup column says whether its target list exists in the target', () => {
  const { source, target } = sites();
  const req = byRel(compareSites(source, target))['Lists/crRequests'];
  assert.equal(req.columns.find((c) => c.name === 'crCustomer').note, 'Lookup to Lists/crCustomers (missing in target)');

  target.allLists.push({ id: 't-cust', title: 'Customers', url: '/sites/prod/Lists/crCustomers' });
  const again = byRel(compareSites(source, target))['Lists/crRequests'];
  assert.equal(again.columns.find((c) => c.name === 'crCustomer').note, 'Lookup to Lists/crCustomers');
});

test('lookups to the same list on two sites compare equal despite different GUIDs', () => {
  const look = (guid) => ({ InternalName: 'crL', TypeAsString: 'Lookup', LookupList: guid, LookupField: 'Title' });
  const a = columnProps(look('{AAA}'), (g) => ({ aaa: 'Lists/crCustomers' })[g.replace(/[{}]/g, '').toLowerCase()]);
  const b = columnProps(look('{BBB}'), (g) => ({ bbb: 'Lists/crCustomers' })[g.replace(/[{}]/g, '').toLowerCase()]);
  assert.deepEqual(propDiffs(a, b), []);
  const c = columnProps(look('{CCC}'), () => 'Lists/crOther');
  const [diff] = propDiffs(a, c, new Set(['Lookup']));
  assert.deepEqual([diff.prop, diff.blocked], ['Lookup', true]);
});

test('same internal name with a different field ID is noted', () => {
  const source = side('/sites/dev', 'Dev', [{ id: 'a', title: 'A', rel: 'Lists/A', fields: [field('crX', { Id: '{1}' })] }]);
  const target = side('/sites/prod', 'Prod', [{ id: 'b', title: 'A', rel: 'Lists/A', fields: [field('crX', { Id: '{2}' })] }]);
  const [a] = compareSites(source, target);
  assert.equal(a.columns[0].status, 'same');
  assert.equal(a.columns[0].note, 'Same internal name, different field ID');
});

test('list status and counts follow the custom-columns switch', () => {
  const source = side('/sites/dev', 'Dev', [{ id: 'a', title: 'A', rel: 'Lists/A', fields: [title, { ...created, Title: 'Erstellt' }] }]);
  const target = side('/sites/prod', 'Prod', [{ id: 'b', title: 'A', rel: 'Lists/A', fields: [title, created] }]);
  const lists = compareSites(source, target);
  assert.equal(listStatus(lists[0], { customOnly: true }), 'same');
  assert.equal(listStatus(lists[0], { customOnly: false }), 'changed');
  assert.deepEqual(countStatuses(lists, { customOnly: true }), { missing: 0, changed: 0, only: 0, same: 2, error: 0 });
  assert.deepEqual(countStatuses(lists, { customOnly: false }), { missing: 0, changed: 1, only: 0, same: 2, error: 0 });
});

test('counts and summary line for the mockup example', () => {
  const { source, target } = sites();
  const lists = compareSites(source, target);
  const n = countStatuses(lists, { customOnly: true });
  // lists: crCustomers, crChangeRequests, crContactsNew missing; OldTasks, crContacts only; Approvals error; Requests changed (settings)
  // Requests columns: Budget, Customer missing; Priority, DueDate changed; LegacyId only; Title, Notes same
  assert.deepEqual(n, { missing: 5, changed: 3, only: 3, same: 2, error: 1 });
  assert.equal(columnSummary(byRel(lists)['Lists/crRequests'], { customOnly: true }), '2 missing · 2 changed · 1 only in target');
});

test('filterView: statuses, text filter, errors always shown', () => {
  const { source, target } = sites();
  const lists = compareSites(source, target);
  const all = filterView(lists, { statuses: ['missing', 'changed', 'only'] });
  assert.ok(all.some((v) => v.list.status === 'error'), 'errors are always shown');
  const req = all.find((v) => v.list.rel === 'Lists/crRequests');
  assert.equal(req.showSettings, true);
  assert.deepEqual(req.columns.map((c) => c.name), ['crBudget', 'crCustomer', 'crDueDate', 'crPriority', 'crLegacyId']);

  const onlyMissing = filterView(lists, { statuses: ['missing'] });
  assert.equal(onlyMissing.find((v) => v.list.rel === 'Lists/crRequests').showSettings, false);
  assert.ok(!onlyMissing.some((v) => v.list.status === 'only'));

  const q = filterView(lists, { statuses: ['missing', 'changed', 'only'], query: 'budget' });
  assert.deepEqual(q.map((v) => v.list.rel), ['Lists/crRequests']);
  assert.deepEqual(q[0].columns.map((c) => c.name), ['crBudget']);

  const noSame = filterView(lists, { statuses: ['missing', 'changed', 'only'] });
  assert.ok(!noSame.some((v) => v.columns.some((c) => c.status === 'same')));
  const withSame = filterView(lists, { statuses: ['same'] });
  assert.deepEqual(withSame.find((v) => v.list.rel === 'Lists/crRequests').columns.map((c) => c.name), ['crNotes', 'Title']);
});

test('diff CSV: one row per item or changed property, same format as the export', () => {
  const { source, target } = sites();
  const view = filterView(compareSites(source, target), { statuses: ['missing', 'changed', 'only'] });
  const rows = diffRows(view);
  const prio = rows.filter((r) => r.Name === 'crPriority');
  assert.deepEqual(prio.map((r) => r.Property), ['Required', 'Choices']);
  assert.equal(prio[1].Source, 'Low; High; Urgent');
  const due = rows.find((r) => r.Name === 'crDueDate' && r.Property === 'Type');
  assert.equal(due.Note, "can't be changed in place");
  assert.equal(rows.find((r) => r.Name === 'Lists/crApprovals').Status, 'Not compared');
  assert.equal(rows.find((r) => r.Name === 'Lists/OldTasks').Status, 'Only in target');

  const csv = diffCsv(rows);
  assert.ok(csv.startsWith('﻿' + DIFF_CSV_COLUMNS.join(';') + '\r\n'));
  assert.ok(csv.endsWith('\r\n'));
  assert.ok(csv.includes('"Low; High; Urgent"'), 'values with the separator are quoted');
});

test('diff CSV quotes values that contain the separator and guards formulas', () => {
  const csv = diffCsv([{ Status: 'Changed', Title: '=SUM(A1)', Source: 'Low; High' }]);
  assert.ok(csv.includes("'=SUM(A1)"));
  assert.ok(csv.includes('"Low; High"'));
});

test('diff file name', () => {
  assert.equal(diffFileName('R&D Hub', new Date('2026-10-21T10:00:00Z')), 'R&D-Hub-compare-2026-10-21.csv');
});
