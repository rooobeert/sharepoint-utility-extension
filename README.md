> **Disclaimer 1:** This extension is fully vibe coded with [Claude Code](https://claude.com/claude-code). Review the code and test it on a non-production site before you rely on it.

> **Disclaimer 2:** The extension is not publicly available in the Microsoft Edge Add-ons store. It can only be installed through this private link: [SharePoint Utility on Microsoft Edge Add-ons](https://microsoftedge.microsoft.com/addons/detail/sharepoint-utility/ijbljmohkldogjablkclbkhjfncfnola)

# SharePoint Utility

A Manifest V3 extension for Chrome and Edge with SharePoint tools for site owners. It works as the signed-in user, from a side panel with these tabs:

- **Batch Runner** sends SharePoint REST `$batch` requests that create lists, libraries and columns.
- **Export Architecture** exports a site's lists, columns, views, content types and permissions to a CSV file, or as ready-made batches that rebuild the lists on another site.
- **Navigation** builds the site's Quick Launch or top navigation from its lists, grouped by solution tag, with drag and drop.
- **Permissions** gives several users and groups permission levels on several lists at once, in `$batch` requests.
- **Compare** shows the differences in lists and columns between two open sites, e.g. dev and prod. It only reads.
- **Search Schema** maps crawled properties (list columns) to `Refinable*` managed properties on the site collection, many at once.
- **Headers** adds request headers (default `Accept: application/json`, optionally `Accept-Language`) when you open SharePoint URLs such as `/_api/` in a browser tab. This uses the `declarativeNetRequest` permission and only changes page loads, not the extension's own requests.

The header shows the extension's version (currently 1.16.0).

## Batch Runner


A Manifest V3 extension for Chrome and Edge that sends SharePoint REST `$batch` requests as the signed-in user. It replaces the Power Automate flow that created lists, libraries and columns with "Send an HTTP request to SharePoint". You paste or load the same batch bodies you used in the flow.

All requests are made from inside the SharePoint tab (`chrome.scripting.executeScript`). They are same-origin and use your normal session cookies. There is no app registration and no token, and the background worker makes no network calls.

```
extension/              the extension (this folder is what you load or zip)
  manifest.json
  background.js         opens the side panel on toolbar click
  sidepanel.html/.css   side panel with the two tabs
  sidepanel.js          Batch Runner tab
  export-panel.js       Export Architecture tab and tab switching
  nav-panel.js          Navigation tab
  permissions-panel.js  Permissions tab
  compare-panel.js      Compare tab
  schema-panel.js       Search Schema tab
  request-headers.js    Headers tab
  shared.js             helpers used by all tabs (dialog, tab requests)
  lib/batch.js          batch body parser and builder
  lib/response.js       multipart response parser, classification, Retry-After
  lib/placeholders.js   ID and URL placeholders, list lookup
  lib/runner.js         stage runner (with automatic parts) and the functions injected into the tab
  lib/log.js            text and JSON run log
  lib/export.js         CSV format and row builders for the architecture export
  lib/exporter.js       reads lists, columns, views, content types, permissions
  lib/provision.js      ready-made batches from an export; the "### STAGE:" file format
  lib/navigation.js     list grouping by solution tag, navigation tree, save plan
  lib/navigator.js      reads and writes the Quick Launch / top navigation
  lib/permissions.js    people picker results, permission plan per list, batch bodies
  lib/permission-client.js  reads lists, levels, groups and role assignments; search; apply
  lib/diff.js           compares the lists and columns of two sites; the diff CSV
  lib/schema.js         search schema: row checks, pending changes, the classic page parsers and postback
tests/                  node --test, no dependencies
samples/                synthetic example batch bodies (no real data) used by the tests
```

## Using it

1. Open any page of the target site, e.g. `https://contoso.sharepoint.com/sites/projects/SitePages/Home.aspx`.
2. Click the extension icon to open the side panel.
3. Add stages. Paste a batch body, use **Load** on a stage for a single `.txt` file, **Load batch file…** for a file with several stages (see *Batch files with stages*), or use **Load folder…** to add every `.txt` file in a folder as one stage each, sorted by file name (`Batch1`, `Batch2`, … `Batch10`). **+ Add stage** is below the last stage. To reorder, drag a stage by the grip **⠿** at the front of its row, or focus the grip and press <kbd>↑</kbd>/<kbd>↓</kbd>.
4. **Preview** is required before every run. It checks the target site (see below), parses every stage and looks up the placeholders. It sends no batch, only `GET` requests. Every check starts collapsed; expand a stage to see its errors, warnings, placeholders and operations.
5. **Run** is enabled once Preview has found no errors. Any edit to a stage turns it off until you run Preview again. Run checks the site once more, asks for confirmation (target site, number of operations, warnings) and then runs the stages in order. Keep the SharePoint tab open on the site until the run finishes.
6. **Copy log** copies the results as plain text. **Export JSON** saves them as a file.

### Errors and warnings

Preview reports two kinds of findings:

- **Errors block the run.** They mean the batch body is malformed, so SharePoint would reject it as a whole, or the site check failed. The errors are: an empty stage; no boundary line at the start; the closing `--batch…--` or `--changeset…--` line missing; a changeset without a boundary; a request line that can't be read; a write operation (POST, MERGE, PATCH, PUT, DELETE) outside a changeset; no operations at all.
- **Warnings don't block the run.** Examples are more than 100 operations, or a placeholder that can't be resolved yet. The second is expected when an earlier stage creates the list: placeholders are looked up again right before each stage runs, and a stage whose placeholders still can't be resolved stops without sending anything.

### Line numbers

Each operation knows the line of its request line (`POST https://…`) in the stage text, counted exactly as shown in the editor. Line numbers appear in the Preview and result tables, in error messages and in the text log. Click a line number to jump to that line in the stage and select it.

If SharePoint rejects a whole batch, it doesn't say which operation caused it. The extension then names the likely operations by line. For *"An invalid HTTP method 'MERGE' was detected for a query operation"*, for example, it lists the MERGE operations and their lines and explains what SharePoint expected.

### Find and replace

<kbd>Ctrl</kbd>+<kbd>F</kbd> (<kbd>Cmd</kbd>+<kbd>F</kbd> on Mac) opens a find/replace bar directly above the stage you are editing. Selected text is filled in as the search term.

- All matches are highlighted, and the current one is shown in orange. A counter shows "3 of 12".
- **Aa** turns on match case.
- <kbd>Enter</kbd> goes to the next match and <kbd>Shift</kbd>+<kbd>Enter</kbd> to the previous one.
- In the Replace field, <kbd>Enter</kbd> replaces the current match. **All** replaces every match in the stage.
- Replacements can be undone with <kbd>Ctrl</kbd>+<kbd>Z</kbd>.
- <kbd>Esc</kbd> closes the bar and leaves the current match selected.

### Target site

A run always targets exactly one site, and the batch bodies already name it: every operation URL starts with the site URL (`https://tenant.sharepoint.com/sites/x/_api/...`). So the site URL is **taken from the batch bodies** and the field is read-only. Before Preview and Run, the extension checks in this order and stops with an explanation if any check fails:

1. **Every operation in every stage targets the same site.** If not, the side panel lists each site with the stages and number of operations that use it.
2. **The site URL comes from those operation URLs.** You can't type a different one.
3. **The active tab is a page of the same site collection.** The extension compares `_spPageContextInfo.siteAbsoluteUrl` from the tab, or the `/sites/x` part of the tab's URL if that isn't available.
4. **SharePoint confirms the site.** `GET {site}/_api/web?$select=Title,Url` must return exactly that URL, and `GET {site}/_api/site?$select=Url` must return the same site collection as the tab. SharePoint resolves an unknown path to the nearest parent site, so a misspelled site URL is caught here. The confirmation dialog shows the site's title, URL and site collection.

Only if the batches use relative URLs (`/_api/...`, no host) is the field editable. It is then filled from `_spPageContextInfo.webAbsoluteUrl` of the active tab, and checks 3 and 4 still apply.

### What happens per stage

For each stage the runner:

- resolves placeholders; if any can't be resolved, it stops and sends nothing;
- converts line endings to CRLF and adds the closing batch boundary if it is missing;
- fetches a fresh form digest (`POST /_api/contextinfo`) and posts to `{site}/_api/$batch` with `X-RequestDigest`, `Content-Type: multipart/mixed; boundary=<boundary from the body>` and `Accept: application/json;odata=nometadata`;
- splits the response into one result per operation (method, URL, HTTP status, error message).

Results per operation:

| Result | Meaning |
|---|---|
| OK | 2xx |
| Already exists (warning) | The list or field already exists (English and German SharePoint messages). Re-running a set of batches is safe. |
| Failed | Any other error. The run pauses and asks whether to continue with the next stage. |

When the run pauses, the dialog shows that stage's outcome: the summary, any stage error, and collapsed lists of the failed, already existing and succeeded operations. Each list shows line, method, status, URL and SharePoint's message. **Copy this stage's log** copies the stage's text log before you decide whether to continue or stop.

**Throttling.** If the whole batch request returns 429 or 503, the runner waits for `Retry-After` (10 s if the header is missing) and sends it again, up to 3 times. If single operations inside the batch return 429 or 503, only those operations are sent again, unchanged, up to 3 times. The side panel shows a countdown during each wait.

## Placeholders

Batch bodies can contain placeholders that are filled in on the target site. They work whether the batches come from the Power Automate flow, are written by hand or are generated by AI.

### List and web IDs

| Placeholder | Replaced with |
|---|---|
| `{{ListId:<reference>}}` | the list's GUID, without braces |
| `@{variables('<any name>')?['<reference>']}` | the same; this is the Power Automate form, so exported flow bodies run unchanged |
| `{{WebId}}` | the web's GUID |

A `<reference>` can be any of these:

| Reference | Example | Matched against |
|---|---|---|
| Site-relative URL | `Lists/crRequests`, `crDocumentVault` (a library) | the list's root folder, relative to the target web |
| Title | `Change Requests` | the list title |
| Server-relative URL | `/sites/projects/Lists/crRequests` | the list's root folder |
| Full URL | `https://tenant.sharepoint.com/sites/projects/Lists/crRequests` | the list's root folder |

- Matching ignores case, a trailing slash and `%20` encoding.
- A single word such as `crDocumentVault` is tried as a URL name and as a title. If it matches two different lists, Preview reports it as ambiguous; use the full URL then.
- `CreateFieldAsXml` needs braces around a lookup's list ID, so write them yourself: `List="{{{ListId:Lists/crRequests}}}"` becomes `List="{3f2a1c4e-…}"`.
- Placeholder names are not case-sensitive (`{{listid:…}}` works too).

**Recommended for AI-generated batches:** the site-relative form, `{{ListId:Lists/<name>}}` for lists and `{{ListId:<name>}}` for libraries. It is short and doesn't contain the site.

### Lists created in the same batch

You don't need a separate stage for lookups. A lookup can refer to a list that is created earlier in the same stage. The extension then sends the stage in parts, keeping the original order:

1. Everything up to the first operation whose list doesn't exist yet is sent as one batch.
2. All lists are read again, exactly like the flow's `objListGuids` object (server-relative URL → ID), and the new list's ID is filled in.
3. The next part is sent. This repeats if later operations need lists created further down.

Preview shows this in advance, for example "List "Lists/ckmPortfolio": created in this stage at line 8" and "Sent in 2 parts". The progress line shows each part, and the results list each operation under its original number and line. Each stage's list object is included in the JSON export (`stages[].listGuids`).

The run stops before sending anything more, with an explanation and line numbers, when:

- a list doesn't exist and **no earlier operation in the stage creates it** (in a later stage this is fine: it is looked up again before that stage runs);
- the operation that should have **created the list failed**: the message includes that operation's error;
- a list is **used before it is created** in the same stage: Preview warns about this.

A list creation is recognised as `POST …/_api/web/lists` with a `Title`. The title is matched against the reference's title or its last URL segment.

### Site URLs

| Placeholder | Example value |
|---|---|
| `{{WebUrl}}` | `https://tenant.sharepoint.com/sites/projects` |
| `{{WebServerRelativeUrl}}` | `/sites/projects` (empty for a root site) |
| `{{SiteUrl}}` | the site collection's URL |
| `{{SiteServerRelativeUrl}}` | the site collection's server-relative URL |

None of these has a trailing slash. With them, one batch runs on dev, test and prod:

```
POST {{WebUrl}}/_api/web/getList('{{WebServerRelativeUrl}}/Lists/crRequests')/fields/CreateFieldAsXml HTTP/1.1
```

If the batches name no site at all, because they use these placeholders or relative URLs, the site URL field is editable and filled from the active tab. If any stage contains hardcoded URLs, the site is taken from those as before. The URL placeholders are then filled with that site. The values used are listed under *Placeholders* in Preview and in the results.

## Export Architecture

Open any page of the site, open the side panel and switch to **Export Architecture**. The tab reads the site's lists and shows them in a table with type, item count and whether permissions are unique or inherited.

1. **Pick the lists.** Use the checkboxes, **All**/**None** (they apply to the lists currently shown) or the filter. Nothing is selected at first. System lists (hidden lists, catalogs such as `_catalogs/masterpage`, SharePoint's system lists) appear only with **Show system lists**.
2. **Choose the columns.** By default, only columns added to the list are exported, plus `Title`. **Include built-in and hidden columns** exports every field.
3. **Export CSV** reads the site and each selected list and downloads `<site title>-architecture-<date>.csv`. A summary shows the number of rows per type.

Everything is read with `GET` requests from the SharePoint tab; nothing is changed. If a part can't be read, it becomes a row with "(… could not be read)" and the error, the export continues, and the part is listed under *Parts that could not be read*. The usual cause is permissions you aren't allowed to see. SharePoint throttling (429/503) is retried as in the Batch Runner.

### CSV format

- **Encoding and layout:** UTF-8 with BOM, separator `;` (German Excel opens it directly in columns), CRLF line ends, quoting per RFC 4180. Line breaks inside XML stay inside one quoted cell.
- **Formula guard:** a value that starts with `=`, `+`, `-`, `@`, a tab or a CR gets a leading `'`, so Excel can't run it as a formula.
- **Excel cell limit:** Excel shows at most 32,767 characters per cell. Very large view or content type XML is cut off in Excel but complete in the file.

There is one row per item, and the kind of item is in `Type`. Rows come in this order: Site, SitePermission, then for each list: List, ListPermission, ListColumn, ListView, ListContentType.

| Column | Site | SitePermission / ListPermission | List | ListColumn | ListView | ListContentType |
|---|---|---|---|---|---|---|
| `Type` | Site | SitePermission / ListPermission | List | ListColumn | ListView | ListContentType |
| `List` | | – / list title | list title | list title | list title | list title |
| `ListUrl` | | – / list URL | server-relative URL | ← | ← | ← |
| `Name` | web URL | principal login | root folder URL | **internal name** | view URL | content type ID |
| `Title` | site title | principal display name | list title | display name | view title | content type name |
| `Id` | web GUID | principal ID | list GUID | field GUID | view GUID | content type ID |
| `Kind` | web template (`STS#3`) | User, SharePointGroup, SecurityGroup | template (`GenericList (100)`) | field type (`Choice`, `Lookup`, …) | view type, `(default)` | content type group |
| `Hidden` | | | TRUE/FALSE | TRUE/FALSE | TRUE/FALSE | TRUE/FALSE |
| `Required` | | | | TRUE/FALSE | | |
| `Roles` | | `Full Control; Edit` (hidden roles such as Limited Access left out) | | | | |
| `Description` | description | | description | description | | description |
| `Settings` | JSON | JSON | JSON | JSON | JSON | JSON |
| `SchemaXml` | | | | field `SchemaXml` | `ListViewXml` | content type `SchemaXml` |

What `Settings` contains, as JSON with empty values left out:

- **Site:** URL, server-relative URL, site collection URL, language, created date, unique permissions yes/no.
- **Permissions:** `Inherited` (whether a list's permissions come from the site) and the user's e-mail address.
- **List:** item count, versioning (major/minor and limits), content approval, forced check-out, content types enabled, attachments, folders, `NoCrawl`, unique permissions, default view URL, created date, last modified.
- **ListColumn:** group, default value, indexed, enforce unique values, read-only, max length, choices, multiple values, lookup target (list title and URL instead of only the GUID, plus the lookup field), selection mode, display format, rich text, number of lines, formula.
- **ListView:** default and personal flags, columns in order, row limit, paging, scope, CAML query.
- **ListContentType:** columns (field links), read-only, sealed.

A column's `SchemaXml` can be used directly in a `CreateFieldAsXml` batch for the Batch Runner. To rebuild a lookup on another site, replace the `List="{…}"` GUID with a placeholder such as `{{{ListId:Lists/<name>}}}`.

## Export as batches

**Export batches**, in the Export Architecture tab, turns the selected lists into a `.txt` file with ready-made stages for the Batch Runner. Use it to rebuild the lists on another site: dev → test → prod, or a template site → new project sites. The file doesn't contain the source site's URL, so it runs on whichever site is open.

After the export, **Load into Batch Runner** puts the stages straight into the Batch Runner tab. The downloaded file can be loaded later with **Load batch file…**.

| Stage | What it does |
|---|---|
| 01 Lists and libraries | `POST /_api/web/lists` with template and description. Each list is created under its **URL name** (e.g. `crRequests`), so its URL is the same as on the source site. |
| 02 List settings | `MERGE` versioning (and limits), content approval, content types on/off, folders, `NoCrawl`; for lists attachments, for libraries minor versions and forced check-out. |
| 03 Columns | `CreateFieldAsXml` with each column's `SchemaXml`, **keeping the field ID**. Lookups become `{{{ListId:Lists/<name>}}}`. A changed `Title` column (display name, required) becomes a `MERGE`. |
| 04 Views | The default view is changed in place (query, row limit, paging, columns), whatever its title is in the target site's language. Other views are created under their URL name with their columns, query and row limit. |
| 05 Display names | Lists and views are renamed from their URL name to their display title. |

How the file is built:

- **Site-independent:** every URL uses `{{WebUrl}}` / `{{WebServerRelativeUrl}}`.
- **Stage size:** stages with more than 50 operations are split into parts, e.g. "03 Columns (1/3)".
- **Cleaned field XML:** what ties a column to the source site is removed from its `SchemaXml`: `SourceID`, `WebId`, `Version`, `ColName`, `RowOrdinal`, and the lookup list GUID.
- **Field IDs kept:** the same column has the same ID on every site. Running the file twice against one site reports "already exists" instead of creating duplicates.

Not exported; the file header lists what was skipped:

- **Permissions and content types:** these are in the CSV only. Users and groups have different IDs on each site, and list content types depend on site content types.
- **Managed metadata columns:** they depend on the term store and a hidden companion field. Create them in SharePoint.
- **Built-in, hidden and read-only columns,** including the extra fields of dependent lookups.
- **Hidden and personal views,** and views other than standard and grid (calendar, Gantt, …).
- **Libraries every site already has** (Documents, Site Pages, Site Assets, Style Library): they are not created. Their columns and views are exported, but they aren't renamed.

A good way to check a rebuild: export the CSV of the source site, run the batches on an empty site, export that site's CSV, and compare the two.

### Batch files with stages

A text file can hold several stages:

```
# Comment lines before the first stage are ignored.
### STAGE: 01 Lists and libraries
--batch_provisioning
…
--batch_provisioning--

### STAGE: 03 Columns
--batch_provisioning
…
```

Each `### STAGE: <name>` line starts a stage. The text up to the next marker must be a complete batch body. **Load batch file…** and **Load folder…** both understand this format. A file without markers is loaded as a single stage, as before.

## Navigation

Builds the site's navigation from its lists and libraries instead of copying each link by hand.

**Lists and libraries** groups the site's lists by the solution tag at the start of their internal (URL) name: the lowercase letters before the first capital, digit or underscore. `hrpLocations` and `hrpEmployees` go into **HRP**, `tdasTermDeals` into **TDAS**. A tag needs at least two lists; everything else is under *Other lists and libraries*. The groups start collapsed. The chip shows how many lists of the group are already in the navigation (a link to the list's root folder or a page below it, e.g. a view).

- **+ Add as menu** adds a heading named after the tag (HRP) at the end, or reuses a top-level heading with that name, and puts every list of the group that isn't linked yet below it.
- **+** on a list adds it under its group's heading (created if missing). Lists without a tag are added at the end.
- Or drag a list, or a whole group, into the navigation tree.

Links use the list's title and its root-folder URL. Both can be edited in the tree.

**Navigation** is the tree that will be saved. It starts as the navigation that is on the site now. Drag ⠿ to reorder; dropping on the top or bottom of an item puts it before or after, dropping on the middle nests it. With ⠿ focused, ↑/↓ reorder, → nests under the item above and ← moves out. SharePoint shows at most three levels, so deeper drops are refused. A heading without a link has an empty URL (saved as SharePoint's `http://linkless.header/`).

**Which navigation** is read from the site: communication sites use the top navigation; team sites use the Quick Launch, whether it is shown on the left or, with the horizontal layout, at the top (SharePoint stores both as the Quick Launch). The panel says what it found, also when the site navigation is turned off. For other cases, such as classic sites that show both, a link below lets you edit the other navigation instead. Hub navigation is not changed.

**Save navigation** shows what will be added, renamed, removed and recreated, and sends nothing until you confirm. SharePoint can only add a node at the end of its level and can't move nodes, so a moved node, and everything after the first new or moved node on its level, is deleted and created again. Audience targeting on recreated nodes is lost. Nodes that stay in place keep their ID and settings. After saving, the tree is read again from SharePoint; reload the SharePoint page to see the result. Requests that fail are listed and the rest still run.

## Permissions

Gives several users and groups permission levels on several lists in one go. It replaces the Power App that collected list/principal/level combinations and set them one by one with Power Automate.

1. **Lists.** Pick one or more lists. The table shows whether each list has unique or inherited permissions. **Reload** reads lists, levels and groups again and clears the plan; **Reload levels** only reads the permission levels again (e.g. after you created a custom level on the site) and keeps the plan.
2. **Users and groups.** Type in the search field to search the tenant's users, security groups and Microsoft 365 groups with SharePoint's own people picker (`clientPeoplePickerSearchUser`, the one behind *Share* and *Grant permissions*). No Graph and no app registration is needed. Clicking the empty field lists the site's SharePoint groups, except the default groups (Owners, Members, Visitors), which are handled by *Keep these default groups*. Typing filters that list and also searches the tenant (up to 50 matches, so a name prefix such as `M365_SOL_HR` lists that whole family). Every entry is labelled *SPO Group*, *Entra Group* (security or Microsoft 365 group) or *User*. Each picked principal gets its own permission level; *Read* is preselected.
   **Item-level permissions (lists only)** sets the list settings page's *Read access* (all items / items created by the user) and *Create and Edit access* (all items / own items / none). Leave a field on *Don't change* to keep it. Libraries don't have this setting and are skipped.
3. **Add to plan** adds every selected list × every picked principal with its level, and the item-level setting if one is chosen. You can add an item-level setting without users or groups; such a list's permissions are then left alone, also in replace mode. Repeat with other lists or people. The plan is grouped by list; single entries or whole lists can be removed.
4. **Current permissions of the lists** applies to the whole run:
   - *Keep them and add the planned ones*: nothing is removed. A level the principal already has is skipped.
   - *Replace them with the planned ones*: every other user and group is removed from the list. Under *Keep these default groups* you tick which of the site's Owners, Members and Visitors groups stay; the others are removed like everyone else. Owners is ticked at first; the checkboxes show the site's actual group names. A planned principal that already has other levels loses those.
5. **Preview** reads each list's current role assignments and shows per list what will be added, removed, kept or skipped, and whether the list stops inheriting. It changes nothing. It warns when the Owners group loses access or nobody keeps Full Control.
6. **Apply** asks for confirmation, then sends the changes.

How it is applied:

- A list that inherits is first given unique permissions with `breakroleinheritance(copyRoleAssignments=true, clearSubscopes=false)`. It starts from a copy of the inherited permissions, and everything that is then removed is listed in Preview. Item and folder permissions below the list are not touched.
- Users and groups that only have *Limited Access* are always kept. SharePoint manages that level for sharing links and item permissions.
- Users and Entra groups that are not yet known to the site are added with `web/ensureuser` (as the SharePoint *Share* dialog does) right before the batch. If that fails, only that principal's additions are skipped.
- Item-level settings are one `MERGE` on the list (`ReadSecurity`, `WriteSecurity`), only for the values that differ. They don't need unique permissions, so they never break inheritance.
- Changes are sent as `addroleassignment` / `removeroleassignment` operations in `$batch` requests of up to 100 operations. A list's operations stay in one batch. Throttling is retried as in the Batch Runner. The results show every operation per list; **Copy log** copies them as text.

## Compare

**Compare** shows how the lists and columns of two sites differ, for example before moving changes from a dev site to prod. It only reads; nothing is changed on either site.

1. Open the **target** site (e.g. prod) in the active tab, and the **source** site (e.g. dev) in another tab of the same browser. The source can be on another tenant.
2. Switch to **Compare**. The target is the site in the active tab; pick the source from the open SharePoint tabs. **Reload tabs** reads the open tabs again. The target also follows when you switch tabs.
3. **Compare** reads the lists and columns of both sites, each in its own tab with that tab's sign-in. System lists are left out unless **Include system lists** is checked. Keep both tabs open until it finishes.

How items are matched:

- **Lists** by their URL within the site (`Lists/crRequests`), not by title. Lists created by *Export batches* keep their URL, so they match.
- **Columns** by internal name within a matched list. Display names may differ, for example on a German and an English site; that is reported as a change.
- **Lookups** by the URL of the list they point to, not its GUID. GUIDs differ between sites, so they aren't compared.
- **Default libraries** (Documents, Site Pages, …) are only compared when both sites have them.

Each list and column gets one status:

| Status | Meaning |
|---|---|
| Missing in target | only on the source |
| Changed | on both, with differences. Each difference reads *target → source*: what the target would become. |
| Only in target | only on the target, e.g. a column that is no longer used |
| Same | on both, no differences (hidden at first) |
| Not compared | a list or its columns could not be read on one side; the reason is shown |

What is compared:

- **List settings:** template, title, description, versioning and its limits, content approval, content types, folders, search, attachments (lists), minor versions and check-out (libraries).
- **Columns:** type, display name, required, description, default value, choices, max length, number of lines, rich text, indexed, unique values, multiple values, lookup list and column, formula.

Some changes are marked **can't be changed in place**: a list's template, and a column's type, lookup list or multiple values. SharePoint can't change these on an existing list or column.

Other hints:

- A missing lookup column says whether its lookup list exists in the target.
- A column with the same internal name but a different field ID is noted.
- A missing list and a target-only list with the same title are pointed out, because they are probably the same list under another URL.

Using the results:

- The chips at the top count lists and columns per status. Their checkboxes choose what is shown; *Not compared* is always shown.
- **Custom columns only** (on at first) uses the same rule as the export: columns added to the list, plus `Title`.
- The filter searches list titles, URLs, column names and internal names.
- **Download CSV** saves what is shown as `<target title>-compare-<date>.csv`, in the same format as the architecture CSV (UTF-8 with BOM, `;`, formula guard). The CSV has one row per list or column, and one row per changed property, with the columns `Status`, `Item`, `List`, `ListUrl`, `Name`, `Title`, `Property`, `Source`, `Target` and `Note`.

Views, content types and permissions are not compared yet.

## Search Schema

Maps crawled properties to managed properties in the site collection's search schema (*Site settings → Search Schema*), for many columns at once. It replaces the `sp-schema-mapper` bookmarklet.

**Unsupported by Microsoft.** There is no REST API for the site collection search schema. The tab reads the classic pages `crawledproperty.aspx`, `managedproperty.aspx` and `listmanagedproperties.aspx` and, on save, sends the same form postback as the OK button on `crawledproperty.aspx`. If Microsoft changes those pages, checks or saving fail with an error on the row; nothing else is touched.

Using it:

1. Open any page of the site collection in the active tab and open **Search Schema**. The site collection follows the active tab.
2. Fill in one row per mapping: a crawled property (`ows_<InternalName>`, `ows_taxId_<InternalName>` for managed metadata, or `name@propertyset`) and a managed property (`RefinableString00` …). Both fields open a list of suggestions that filters as you type (↑/↓, Enter, Esc): the list columns of the site with **Not mapped** or **→ RefinableString07**, and all `Refinable*` properties with **Free** or **← ows_Region** when another crawled property already feeds them. You can also paste rows from Excel into a crawled property field: `crawled ⇥ managed [⇥ managed …]` per line (`;`, `,`, `=>` and `->` work too).
3. Each row is checked: does the crawled property exist and what is it mapped to, does the managed property exist, is it free or fed by other crawled properties, and do the types fit (e.g. a text column into `RefinableInt` gets a warning).
4. Click **×** on a mapping under a crawled property to remove it on save (**↶** undoes).
5. **Save (+n −n)** shows the changes per crawled property, writes them, and reads each crawled property page back to confirm. Reindex the list or library afterwards so the values appear in search.

The mapping status behind the suggestions comes from `listmanagedproperties.aspx` (filtered per `Refinable*` prefix, 50 rows per page). It is cached for an hour while the side panel is open; **Refresh status** reloads it. Reading it takes a while on big sites; until it is done, a line above the rows says so, and the suggestions show names without their mapping. Rows and the status are not stored: they are gone when the side panel closes.

## Loading unpacked

**Chrome:** go to `chrome://extensions`, turn on **Developer mode**, click **Load unpacked** and select the `extension` folder. Pin the extension and click its icon on a SharePoint page.

**Edge:** go to `edge://extensions`, turn on **Developer mode** (bottom left), click **Load unpacked** and select the `extension` folder.

Side panels need Chrome or Edge 116 or later. On managed laptops, IT policy can block developer mode or unpacked extensions (`ExtensionInstallBlocklist`, `DeveloperToolsAvailability`). In that case, have the packaged extension published privately or force-installed by IT.

## Packaging for the Chrome Web Store and Edge Add-ons

Both stores take a `.zip` with `manifest.json` at its root. Build it with:

```bash
npm run package
```

This runs the tests, checks the manifest (version format, description of at most 132 characters, 128×128 icon present) and writes `dist/sharepoint-utility-<version>.zip`. `dist/` is not part of the source. The icons in `extension/icons/` and the store logo in `store/` are drawn by `python3 scripts/draw-icon.py` (needs Pillow).

Then upload the zip:

- **Edge Add-ons:** Partner Center → Microsoft Edge → Create new extension. The listing text and answers for every page cover: visibility (*Hidden* for internal use), permission justifications, data usage, the long description, the 300×300 logo in `store/`, and the certification notes. IT can also force-install the published extension through the `ExtensionInstallForcelist` policy.
- **Chrome Web Store:** Developer Dashboard → New item. For company-only use, set visibility to *Private* (your Google Workspace domain) or *Unlisted*. The same listing text works.

For every update, raise `version` in `manifest.json` (patch for fixes, minor for features), run `npm run package` again and upload the new zip. The side panel header shows the version.

## Permissions and data

- `host_permissions`: `https://*.sharepoint.com/*` only.
- `permissions`: `scripting`, `sidePanel`, `storage`.
- Nothing is sent anywhere except the SharePoint site. There are no external libraries, no CDN and no analytics.
- `chrome.storage.local` holds only the last site URL and the stage names and bodies. The export, the comparison, the navigation builder, the permissions tool and the search schema tool store nothing (the permissions plan lives only in the open side panel); the CSV is created in the side panel and downloaded directly.
- Cookies are never read. The form digest is fetched and used inside the page and never returned to the extension, logged or stored. Run logs contain methods, URLs, status codes and error messages, but no request bodies.

## Known limits

- **Batch size.** SharePoint Online rejects a batch with more than 100 operations, and large batches are throttled more often. Keep each stage to a few dozen operations. Preview and the confirmation dialog warn above 100. In `samples/`, `Batch3.txt` (132) and `Batch5-other-tenant.txt` (108) must be split before they will run.
- **Digest lifetime.** A form digest is valid for about 30 minutes (the site's `FormDigestTimeoutSeconds`). The runner fetches a new one before every stage and every retry, so only one very slow request could run into it.
- **No rollback.** SharePoint does not treat a changeset as a transaction. Operations before and after a failure still run, so fix the cause and re-run the whole stage. Operations that already succeeded come back as "already exists".
- **One site per run.** All stages must target the same site (see *Target site*). For example, `Batch5-other-tenant.txt` targets `https://fabrikam.sharepoint.com/sites/ops-use-cases`, which is a different tenant, so it can't be run together with the `rd-hub` batches.
- **Permissions on the site.** The user needs to be a site owner, or have *Manage Lists* plus *Edit*, on the target site. Display-name updates with `Accept-Language` only take effect for languages enabled on the site.
- **The tab must stay open.** Requests run in the SharePoint tab. If you navigate that tab to another tenant or close it during a run, the current stage stops with an error.
- **Placeholder lookups** find lists by title or root-folder URL in the target web only. A URL pointing inside a list (a view or a folder) is not matched; use the list's root URL, as the flow does.
- **Automatic parts** recognise list creation only through `POST …/_api/web/lists`. If you create lists another way, put the lookups in a later stage.

## Tests

```bash
node --test
```

Node 18 or later, no `npm install` needed. `package.json` exists only to mark the files as ES modules and is not part of the extension. The tests cover:

- the batch parser against every file in `samples/`: operation counts, CRLF normalisation, closing boundary, rebuilding a subset;
- the response parser against `tests/fixtures/response-mixed.txt`: success, duplicate field, 429 with `Retry-After`, a real failure and a 204;
- placeholder discovery, the list object built from one list read, and replacement, including the Power Automate alias;
- all reference forms (title, site-relative, server-relative, full URL, Power Automate with any variable name), ambiguity, and the URL placeholders;
- automatic parts with a stateful fake SharePoint: one and two dependency levels, a failed list creation, a list nothing creates, a library by URL name, a rejected second part;
- deriving the one target site from the batch bodies and reporting conflicts;
- the stage runner with a fake page: site verification, placeholder stop, inner-429 retry, batch-level 503, digest failure, pausing between stages, and Preview sending nothing;
- the text and JSON log;
- the permissions tool: people picker results, permission levels, site groups, the plan in add and replace mode (default groups, Limited Access, warnings), item-level settings, the batch body, packing into batches of 100, and applying with a fake page;
- the navigation builder: solution tags and grouping, link matching, the save plan (keep, rename, add, remove, recreate after a move) and saving with a fake page;
- the batch export: cleaned field XML, stage order and contents, URL encoding, splitting large stages, reading the file back, and a full round trip through the Batch Runner on an empty fake site;
- the search schema tool: crawled names from list columns, type warnings, the mapping status, paste parsing, row states, pending changes per crawled property, and the mappings to post and verify;
- the comparison: matching lists by URL across site collections, column statuses, choices added and removed, lookups by target URL, changes that can't be made in place, unreadable lists, default libraries, same-title hints, counts and filters with and without custom columns, and the diff CSV;
- the architecture export: CSV quoting, BOM, separator, formula guard, every row type, custom-column filter, lookup targets, hidden roles, and the exporter with a fake SharePoint (an unreadable part, 429 retry, fallback when `IsSystemList` is unknown).
