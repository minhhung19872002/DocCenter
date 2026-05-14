# Document Center — SPFx Web Part

SharePoint Framework web part for SharePoint Online that lets users:

- **Upload** documents into a document library, tagging each with one or more hashtags.
- **Search** documents by clicking hashtag pills (AND/OR mode) and optional file name filter.
- **Admin** tab (visible only to Site Owners) for managing the hashtag list — add, rename, delete.

Hashtags are stored as a separate SharePoint **list**, and the document library has a multi-lookup field pointing at it. Renaming a hashtag updates it everywhere automatically.

---

## Prerequisites

- Node.js **18.x** (LTS)
- gulp CLI (`npm i -g gulp-cli`)
- Yeoman + SPFx generator (only if you plan to scaffold related projects)
- A SharePoint Online site where you have Owner or Site Collection Admin rights
- A SharePoint **App Catalog** for the tenant (or a site-collection app catalog)

## First-time setup

```powershell
cd C:\Source\DocCenter
npm install
```

If you hit SSL trust errors on `gulp serve`, run once:

```powershell
gulp trust-dev-cert
```

## Local development (workbench)

1. Open `config/serve.json` and change `pageUrl` to a real SharePoint page on your dev site:
   `https://<tenant>.sharepoint.com/sites/<yoursite>/SitePages/Work.aspx`
2. Start the dev server:
   ```powershell
   gulp serve
   ```
3. In the page that opens, edit it, add the **Document Center** web part. On first load it will create the `Documents` library and `Hashtags` list automatically.

## Production build & deploy

```powershell
gulp clean
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/doc-center.sppkg`. Upload it to your tenant App Catalog (or site-collection app catalog) and click **Deploy**. Then add the **Document Center** app to the destination site and drop the web part on any page.

## Web part properties

Open the web part's property pane to override:

- **Documents library title** — defaults to `Documents`
- **Hashtags list title** — defaults to `Hashtags`

If the library/list does not exist, it is created on first render.

## How hashtags are stored

- A custom list **Hashtags** (template 100) — one item per hashtag, `Title` is the hashtag text.
- A column **DocHashtags** (display name **Hashtags**) is added to the documents library as a **multi-value Lookup** pointing at the Hashtags list.
- Search reads documents and expands the lookup field, then filters by selected hashtag IDs.

## Admin access

The Admin tab is visible only to members of the site's **Owners** group (the group returned by `web.associatedOwnerGroup`). To grant admin rights to another user, simply add them to the Site Owners group in **Site Settings → Site Permissions**.

## Project layout

```
src/webparts/docCenter/
  DocCenterWebPart.ts            — web part entry, property pane
  DocCenterWebPart.manifest.json — manifest
  components/
    DocCenter.tsx                — root + tabs
    Upload.tsx                   — drag/drop upload + hashtag picker
    Search.tsx                   — hashtag-driven search
    Admin.tsx                    — hashtag CRUD
    DocCenter.module.scss
    IDocCenterProps.ts
  services/
    SharePointService.ts         — PnP JS: provision, CRUD, search, admin check
    types.ts
  loc/
    en-us.js, mystrings.d.ts
```

## Notes

- Search currently fetches up to 5000 items from the library and filters client-side; for larger libraries swap in the SharePoint Search REST API (`/_api/search/query`) instead.
- File upload uses `addUsingPath` (chunked) — works for files up to 250 MB. For larger files, switch to `addChunked`.
- Multi-lookup field is created via `createFieldAsXml`; if you rename the field afterwards, update `HASHTAG_FIELD_INTERNAL` in `SharePointService.ts`.
