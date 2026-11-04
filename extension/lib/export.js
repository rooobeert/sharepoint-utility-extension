// "Export SharePoint Architecture": turns REST responses into CSV rows.
// Pure functions: no DOM, no chrome.* APIs.
//
// One CSV, one row per item, the item's kind in `Type`:
//   Site, SitePermission, List, ListPermission, ListColumn, ListView, ListContentType

export const CSV_COLUMNS = [
  'Type', 'List', 'ListUrl', 'Name', 'Title', 'Id', 'Kind',
  'Hidden', 'Required', 'Roles', 'Description', 'Settings', 'SchemaXml',
];

const TEMPLATES = {
  100: 'GenericList', 101: 'DocumentLibrary', 102: 'Survey', 103: 'Links', 104: 'Announcements',
  105: 'Contacts', 106: 'Events', 107: 'Tasks', 108: 'DiscussionBoard', 109: 'PictureLibrary',
  110: 'DataSources', 115: 'XMLForm', 119: 'WebPageLibrary', 150: 'GanttTasks', 171: 'TasksWithTimelineAndHierarchy',
  850: 'PublishingPages', 851: 'AssetLibrary', 1100: 'IssueTracking', 3100: 'AccessApp',
};
const PRINCIPAL_TYPES = { 1: 'User', 2: 'DistributionList', 4: 'SecurityGroup', 8: 'SharePointGroup', 15: 'All' };

export const templateName = (t) => (TEMPLATES[t] ? `${TEMPLATES[t]} (${t})` : String(t ?? ''));
export const principalTypeName = (t) => PRINCIPAL_TYPES[t] || String(t ?? '');

const bool = (v) => (v === true ? 'TRUE' : v === false ? 'FALSE' : '');
const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];
const items = (v) => (Array.isArray(v) ? v : (v && (v.Items || v.results)) || []);
const stripBraces = (g) => String(g || '').replace(/^\{|\}$/g, '').toLowerCase();
const clean = (o) => JSON.stringify(Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== null && v !== '')));

// ---------------------------------------------------------------- CSV

/**
 * One CSV cell: RFC 4180 quoting, and a leading apostrophe for values Excel
 * would run as a formula (=, +, -, @, tab, CR) so a list title like
 * "=HYPERLINK(...)" can't execute when the file is opened.
 */
export function csvCell(value, sep) {
  let s = value == null ? '' : String(value);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return /["\r\n]/.test(s) || s.includes(sep) ? `"${s.replace(/"/g, '""')}"` : s;
}

/** CSV text with BOM (Excel then reads UTF-8 correctly) and CRLF line ends. */
export function toCsv(rows, sep = ';') {
  const lines = [CSV_COLUMNS.map((c) => csvCell(c, sep)).join(sep)];
  for (const r of rows) lines.push(CSV_COLUMNS.map((c) => csvCell(r[c], sep)).join(sep));
  return '﻿' + lines.join('\r\n') + '\r\n';
}

export function csvFileName(siteTitle, date = new Date()) {
  const safe = String(siteTitle || 'site').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, '-').slice(0, 60) || 'site';
  return `${safe}-architecture-${date.toISOString().slice(0, 10)}.csv`;
}

// ---------------------------------------------------------------- lists for the picker

/** Hidden lists, catalogs (_catalogs/...) and SharePoint's system lists. */
export function isSystemList(list) {
  const url = (list.RootFolder && list.RootFolder.ServerRelativeUrl) || '';
  return !!(list.Hidden || list.IsCatalog || list.IsSystemList || /\/_catalogs\//i.test(url));
}

/** Compact list info for the picker, sorted by title. */
export function pickerLists(json) {
  return results(json)
    .map((l) => ({
      id: stripBraces(l.Id),
      title: l.Title,
      url: (l.RootFolder && l.RootFolder.ServerRelativeUrl) || '',
      template: l.BaseTemplate,
      baseType: l.BaseType,
      kind: templateName(l.BaseTemplate),
      itemCount: l.ItemCount,
      uniquePermissions: l.HasUniqueRoleAssignments,
      system: isSystemList(l),
    }))
    .sort((a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
}

// ---------------------------------------------------------------- rows

export function siteRow(web, siteCollectionUrl) {
  return {
    Type: 'Site',
    Name: web.Url,
    Title: web.Title,
    Id: stripBraces(web.Id),
    Kind: web.WebTemplate ? `${web.WebTemplate}#${web.Configuration ?? 0}` : '',
    Description: web.Description,
    Settings: clean({
      Url: web.Url,
      ServerRelativeUrl: web.ServerRelativeUrl,
      SiteCollectionUrl: siteCollectionUrl,
      Language: web.Language,
      Created: web.Created,
      HasUniqueRoleAssignments: web.HasUniqueRoleAssignments,
    }),
  };
}

/**
 * Permission rows from roleassignments?$expand=Member,RoleDefinitionBindings.
 * Hidden roles ("Limited Access") are left out; principals with only hidden
 * roles get no row.
 */
export function permissionRows(json, { type, list, inherited }) {
  return results(json)
    .map((ra) => {
      const roles = items(ra.RoleDefinitionBindings).filter((r) => !r.Hidden).map((r) => r.Name);
      if (!roles.length) return null;
      const m = ra.Member || {};
      return {
        Type: type,
        List: list && list.Title,
        ListUrl: list && list.RootFolder && list.RootFolder.ServerRelativeUrl,
        Name: m.LoginName,
        Title: m.Title,
        Id: m.Id ?? ra.PrincipalId,
        Kind: principalTypeName(m.PrincipalType),
        Roles: roles.join('; '),
        Settings: clean({ Inherited: inherited, Email: m.Email || undefined }),
      };
    })
    .filter(Boolean);
}

export function listRow(list) {
  const url = list.RootFolder && list.RootFolder.ServerRelativeUrl;
  return {
    Type: 'List',
    List: list.Title,
    ListUrl: url,
    Name: url,
    Title: list.Title,
    Id: stripBraces(list.Id),
    Kind: templateName(list.BaseTemplate),
    Hidden: bool(list.Hidden),
    Description: list.Description,
    Settings: clean({
      ItemCount: list.ItemCount,
      EnableVersioning: list.EnableVersioning,
      MajorVersionLimit: list.EnableVersioning ? list.MajorVersionLimit : undefined,
      EnableMinorVersions: list.EnableMinorVersions || undefined,
      MajorWithMinorVersionsLimit: list.EnableMinorVersions ? list.MajorWithMinorVersionsLimit : undefined,
      EnableModeration: list.EnableModeration,
      ForceCheckout: list.ForceCheckout,
      ContentTypesEnabled: list.ContentTypesEnabled,
      EnableAttachments: list.EnableAttachments,
      EnableFolderCreation: list.EnableFolderCreation,
      NoCrawl: list.NoCrawl,
      HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
      DefaultViewUrl: list.DefaultView && list.DefaultView.ServerRelativeUrl,
      Created: list.Created,
      LastItemModifiedDate: list.LastItemModifiedDate,
    }),
  };
}

/** Columns added by the user (plus Title), as opposed to SharePoint's built-in and hidden ones. */
export function isCustomField(f) {
  if (f.InternalName === 'Title') return true;
  return !f.FromBaseType && !f.Hidden && !f.ReadOnlyField && !f.Sealed;
}

/**
 * Column rows. `listsById` maps list GUIDs (lower case, no braces) to
 * { title, url } so a lookup's target list is readable, not just a GUID.
 */
export function columnRows(list, fieldsJson, { allColumns = false, listsById = {} } = {}) {
  const base = { Type: 'ListColumn', List: list.Title, ListUrl: list.RootFolder && list.RootFolder.ServerRelativeUrl };
  return results(fieldsJson)
    .filter((f) => allColumns || isCustomField(f))
    .map((f) => {
      const target = f.LookupList ? listsById[stripBraces(f.LookupList)] : null;
      return {
        ...base,
        Name: f.InternalName,
        Title: f.Title,
        Id: stripBraces(f.Id),
        Kind: f.TypeAsString,
        Hidden: bool(f.Hidden),
        Required: bool(f.Required),
        Description: f.Description,
        Settings: clean({
          StaticName: f.StaticName !== f.InternalName ? f.StaticName : undefined,
          Group: f.Group,
          DefaultValue: f.DefaultValue,
          Indexed: f.Indexed || undefined,
          EnforceUniqueValues: f.EnforceUniqueValues || undefined,
          ReadOnlyField: f.ReadOnlyField || undefined,
          FromBaseType: f.FromBaseType || undefined,
          MaxLength: f.MaxLength,
          Choices: f.Choices && items(f.Choices).length ? items(f.Choices) : undefined,
          FillInChoice: f.FillInChoice || undefined,
          AllowMultipleValues: f.AllowMultipleValues || undefined,
          LookupList: f.LookupList ? (target ? target.title : f.LookupList) : undefined,
          LookupListUrl: target ? target.url : undefined,
          LookupField: f.LookupField,
          SelectionMode: f.SelectionMode,
          DisplayFormat: f.DisplayFormat,
          RichText: f.RichText || undefined,
          NumberOfLines: f.NumberOfLines,
          Formula: f.Formula,
        }),
        SchemaXml: f.SchemaXml,
      };
    });
}

export function viewRows(list, viewsJson) {
  const base = { Type: 'ListView', List: list.Title, ListUrl: list.RootFolder && list.RootFolder.ServerRelativeUrl };
  return results(viewsJson).map((v) => ({
    ...base,
    Name: v.ServerRelativeUrl,
    Title: v.Title,
    Id: stripBraces(v.Id),
    Kind: (v.ViewType || '') + (v.DefaultView ? ' (default)' : ''),
    Hidden: bool(v.Hidden),
    Settings: clean({
      DefaultView: v.DefaultView || undefined,
      PersonalView: v.PersonalView || undefined,
      ViewFields: items(v.ViewFields),
      RowLimit: v.RowLimit,
      Paged: v.Paged,
      Scope: v.Scope,
      ViewQuery: v.ViewQuery,
    }),
    SchemaXml: v.ListViewXml,
  }));
}

export function contentTypeRows(list, ctJson) {
  const base = { Type: 'ListContentType', List: list.Title, ListUrl: list.RootFolder && list.RootFolder.ServerRelativeUrl };
  return results(ctJson).map((ct) => {
    const id = (ct.StringId || (ct.Id && ct.Id.StringValue) || '').toString();
    return {
      ...base,
      Name: id,
      Title: ct.Name,
      Id: id,
      Kind: ct.Group,
      Hidden: bool(ct.Hidden),
      Description: ct.Description,
      Settings: clean({
        Fields: items(ct.FieldLinks).map((l) => l.Name),
        ReadOnly: ct.ReadOnly || undefined,
        Sealed: ct.Sealed || undefined,
      }),
      SchemaXml: ct.SchemaXml,
    };
  });
}

/** A row that records a part of the export that could not be read. */
export function errorRow(type, list, what, message) {
  return {
    Type: type,
    List: list && list.Title,
    ListUrl: list && list.RootFolder && list.RootFolder.ServerRelativeUrl,
    Title: `(${what} could not be read)`,
    Description: message,
  };
}
