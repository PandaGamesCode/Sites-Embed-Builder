# Single Site ⚡

A complete Figma-style website builder that runs **entirely inside one Google Sheet + Apps Script**, designed to be embedded as a content page in Google Sites. Two files: `Code.gs` (backend) and `index.html` (the whole app).

## What you get

| Requirement | Where it lives |
|---|---|
| Google sign-in (grabs account email, stores data under it) + username/password login for everyone, all users equal | `googleLogin / registerUser / passwordLogin`, Users tab |
| Optional domain lock (`student.dodea.edu`, company domain) | set script property `ALLOWED_DOMAINS` — empty = allow any Google account |
| 8-hour sessions | `SESSION_HOURS = 8` |
| Auto-created sheet tabs, no quota worries (few users) | Users, Projects, Pages, Versions, Media, Assets, SiteData — created on first run |
| Page-as-JSON-blob storage + version history (last 20 per page, restore anytime) | Pages.dataJson, Versions tab |
| Image/media uploads (small → inline in sheet, large → Drive folder per user) | Media/Assets tabs, `uploadImage` |
| No top nav — pure content page UI; users can still *add* an in-page links bar block if they want | editor shell is headerless content-only |
| Simple / Intermediate / Advanced modes with progressively more options | mode selector; palette & property panel grow with mode; advanced-only features stripped from exports in lower modes |
| Auto-responsive UI built-in (user never writes media queries) | export CSS uses clamp/auto-fit/grid collapse at 760px |
| Live preview while editing | iframe canvas, click-to-select, drag-drop layers like Figma |
| Global styles per project (colors, fonts, radius, dark mode, entrance animations, head/footer code) | Theme modal → Projects.themeJson |
| Export = ONE self-contained HTML file; Copy-all-code button sized for >12k chars | Export modal → `exportPage` |
| Static export = zero server calls (frozen snapshot embedded) | `data-store` attribute baked in |
| Live export = auto-updating site (A/B day, weekly changes) reading the mini-database | LIVE DATA block + `?action=data` endpoint |
| Historic days = buttons rendered from stored day keys showing frozen data | history picker (advanced), freeze-date export |
| Form submissions collected into the sheet from live sites | `?action=form` endpoint |
| Execute as me, department access | deploy settings (see below) |

## One-time deployment (~5 minutes)

1. Create a new Google Sheet → **Extensions ▸ Apps Script**.
2. Replace `Code.gs` with this folder's `Code.gs`; add an HTML file named `index` and paste `index.html`.
3. **Deploy ▸ New deployment ▸ Web app**
   - Execute as: **Me**
   - Who has access: **Anyone in \<your district/department\>** (or Anyone with the link)
4. Open the web-app URL once and authorize scopes (Sheets, Drive, script properties). All tabs auto-create.
5. *(Optional)* lock sign-in to school/company domains: Project Settings ▸ Script Properties ▸ add
   `ALLOWED_DOMAINS = student.dodea.edu,yourcompany.com`
6. Embed the app itself anywhere via **Google Sites ▸ Embed ▸ By URL** (iframe embedding is allowed).

## Publishing a built page to Google Sites

Build pages → **Export** → choose:

- **Static** – fully offline single file (snapshot of today's data embedded).
- **Live ⚡** – same file, but LIVE DATA blocks fetch your sheet-backed mini-database on every visit, so "This week: A Day" updates automatically when you edit the Data modal.
- **Freeze date** (Advanced) – generate a historical copy pinned to any past day; with the *history picker* enabled, visitors get clickable buttons for each stored day.

Then **Copy all code** (works well beyond 12,000 characters) → Google Sites page → Insert ▸ Embed ▸ **Embed code** ▸ paste ▸ Next ▸ Insert. Done — animated, responsive, self-sustaining.

## The mini-database (A/B day example)

Open 🗄 **Data** in the editor. Keys are free-form JSON arrays:

```
schedule = {
  "current": "2026-10-19",
  "2026-10-19": [ {"title":"A Day","detail":"Assembly 9am"} ],
  "2026-10-18": [ {"title":"B Day","detail":"Early release"} ],
  "weeks": { "W42": [ {"title":"Reading Week"} ] }
}
```

A **LIVE DATA** block bound to `current` shows today's entry automatically; ISO-week keys (`W42`) match "this week"; every `YYYY-MM-DD` key becomes a clickable history button in advanced exports. Update the sheet data → every live embed refreshes within seconds. No quotas matter at your scale.

## Editor shortcuts

`Ctrl+S` save (auto-versions) · `Ctrl+Z/Shift+Z` undo/redo · `Ctrl+D` duplicate · `Del` remove · drag blocks in the Layers pane or drop from the palette onto any position · click a block in the live preview to select it.

## Testing

Backend logic (auth, isolation, versions, exports, endpoints) is covered by the harness used during development — 14 end-to-end scenarios against a simulated Spreadsheet/Properties/Drive environment, plus jsdom tests exercising all 28 block types through the renderer, property panels, and preview builder.
