# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A collection of **single-file, zero-dependency SharePoint utilities** for a locked-down GCC / GCC High environment. There is no build system, no package manager, no test runner, and no external libraries, fonts, or CDN calls — every artifact is self-contained and hand-editable. "Building" means editing one file; "running" means opening it in a browser or pasting it into DevTools.

As of this writing the repo contains only `CLAUDE.md` and `SP.code-workspace`; the deliverables below are authored on request and are the actual product. When you create one, verify JavaScript syntax with `node --check <file>` before handing it off (this is the only "test" available).

## The three artifacts and how each runs

1. **SharePoint Console** (standalone `.html`) — a launcher homepage. Paste one or many SharePoint list-view URLs; it parses each into a card grouped by site collection, with Open / Copy / Remove / inline-rename, search, type chips, and JSON export/import. Runs entirely client-side, **no network calls**. Open the local file in a browser or set it as the browser homepage.

2. **SharePoint List Toolkit** (injected panel) — a draggable overlay that reads and writes the current list via the SharePoint REST API (`/_api/...`). Tabs: Overview, Columns, Views, Recycle Bin, Export, Versions.

3. **SharePoint Console Plus** (`.js`, the merged tool) — the Toolkit with the Console folded in as a first **Links** tab. This is the canonical combined deliverable.

### The Gulp tab (crawl → VS Code table)

The snippet's **Gulp** tab crawls a site URL you paste (defaults to `DEFAULT_BASE`, the WHD `usdol.sharepoint.com/sites/whd/mw/mwplanning/` site). Four actions:
- **Gulp lists** — every non-hidden list + library via `<base>/_api/web/lists`; each row is clickable to gulp that list's items.
- **This open list** — gulps the items of the list currently open in the browser, straight from `window._spPageContextInfo` (`ctx.listId`); the button disables itself when the page isn't a list.
- **All lists' items** — walks every list sequentially and pulls each one's items into a single **sectioned bundle** (`lastGulp.sections`), with a summary table of list → row count.
- A single list/items gulp exports **`.md`** (native VS Code Markdown preview, `Ctrl+Shift+V`), **`.json`**, **`.csv`**, and a self-contained sortable/filterable **`.html`**; the multi-list bundle exports `.md` (one `##` section per list), `.json` (object keyed by list title → item rows), and `.html` (section per list) — **CSV is disabled for the bundle** since schemas differ per list.

**Set export folder** (Gulp tab, File System Access API — Edge/Chrome only) redirects **every** export in the panel, all tabs, from the Downloads flow into a user-picked folder (e.g. `C:\SP\gulps`), writing through the shared `downloadBlob()` choke point; with the exported file open in VS Code Live Preview, each gulp refreshes the editor view seconds after the click. The grant lasts one page-load (browser rule — re-pick after refresh), a failed write falls back to a normal download and clears the folder, and unsupported browsers keep the plain download path.

Every discovered list URL is also written into the Links catalog (`saveLinks()`), so crawling feeds the persistent JSON store. The single- and multi-list HTML exports share one `htmlDoc()` shell (`escHtml` / `tableFragment` helpers). Cross-site reads work only within the **same tenant host** the snippet is running on — same-origin session-cookie auth, not headless. `test/gulp.test.js` boots the snippet in jsdom with mocked REST (including a stub `showDirectoryPicker`) and asserts all of this (45 checks); run `cd test && npm i && node gulp.test.js`.

## The one constraint that dictates the architecture

The launcher *can* run as a standalone file, but the REST toolkit **cannot**. REST calls to `/_api/...` authenticate via the user's session cookies, which browsers only permit **same-origin, inside the SharePoint page itself**. Therefore the merged tool ships as a **DevTools Snippet** (or console paste), not an HTML file. A loader bookmarklet only works where the GCC CSP allows remote script, so the Snippet is the reliable path. On a non-SharePoint page the panel still opens but disables the REST tabs, leaving only Links active. Do not propose an architecture that ignores this — it is a browser security fact, not a preference.

## Architecture conventions (keep these consistent across artifacts)

- **Single IIFE**, no globals leaked. The combined tool guards against double-injection by parking a toggle on `window.__SPCP`; re-running toggles instead of rebuilding.
- **Shared storage key** `sp_link_console_v1` in `localStorage`. Both the standalone HTML and the Links tab read/write this same key, so an export from one imports cleanly into the other. Wrap every `localStorage` access in try/catch — it throws in sandboxed contexts (the claude.ai preview blocks it; a downloaded file does not). Reflect failure in a visible status dot rather than failing silently.
- **State is one object**, e.g. `var state = { items: [], q: "", type: "ALL" }`. `items` is the source of truth; every mutation calls `save()` then `render()`. Rendering is full-redraw (data sets are small; this kills stale-DOM bugs).
- **Escape all user text** with an `esc()` helper (`& < > "`) before inserting as HTML — URLs and labels must not inject markup.
- **UI in a Shadow DOM host** for the in-page panel, with all CSS in one `<style>` in the shadow root, so SharePoint's CSS cannot break the panel and vice versa.

### The URL parser (the core of the launcher)

Wrap `new URL()` in try/catch. Split `pathname`, keep a lowercased copy for matching. Find the site collection as the segment after `sites` or `teams` (fall back to host). Detect type by scanning path markers:

| URL contains | Type | Name from |
|---|---|---|
| `/_layouts/` | SYSTEM | last `.aspx` file |
| `/Lists/<name>/AllItems.aspx` | LIST | segment after `Lists` |
| `/Lists/<name>/<other>.aspx` | VIEW | list name + view file |
| `/<library>/Forms/...` | LIBRARY | segment before `Forms` |
| `/SitePages/<page>.aspx` | PAGE | page file |
| any other `.aspx` | PAGE | file name |
| anything else | LINK | last segment or host |

Humanize segments with `pretty()` (decodeURIComponent, then `-`/`_` → space) and `titleCase()`. Detection keys off **path structure, not domain**, so it already handles `.sharepoint.com`, `.sharepoint.us`, and other clouds without new logic.

### REST helpers (the toolkit)

- `spGet` — fetch with `odata=nometadata` to get clean `{ value: [...] }`.
- `spGetAll` — follow `odata.nextLink` to page large collections.
- `getDigest` — POST `/_api/contextinfo` for the form digest required on every write.
- `spPost` — attaches the digest; for updates flips to `X-HTTP-Method: MERGE` with `IF-MATCH: *`.
- Context comes from `window._spPageContextInfo` (web URL, list id), falling back to parsing `/sites/` or `/teams/` from the address; this sets a `restOk` flag gating the REST tabs.

## Safety rails (non-negotiable)

- Every write **confirms first**. Recycle Bin restore and version re-apply both prompt before acting.
- Version re-apply writes back **only** simple text, number, date, choice, and yes/no fields; it skips lookups, people, managed metadata, and attachments (those cannot be reapplied safely).
- Treat Export as the user's rollback snapshot — offer it before any bulk change.
- The tool only ever sees and changes what the user's own SharePoint permissions already allow; the form digest is SharePoint's own protection, obtained legitimately. Do not add anything that circumvents permissions or the digest.

## Visual design — TrueCrimeDev

Both artifacts are skinned with the **TrueCrimeDev** design system (imported from the user's claude.ai/design project via the `DesignSync` tool + `/design-login`). The portable stylesheet lives in **`theme.css`**, namespaced under `.tc`, dark-mode-native with a `.tc-light` remap. Reconstructed from the published spec — if you re-authorize `DesignSync`, prefer swapping in the exact upstream `theme.css` over editing this copy.

Non-negotiable brand rules when touching any UI here:
- **Depth comes from borders, not shadows.** Surfaces lift by getting a lighter background + lighter border. Shadows are reserved for genuinely floating layers (toast, dropdowns).
- **95% greyscale.** The six accents (`--tc-blue/green/amber/red/purple/cyan`) carry meaning only — status, primary actions, links. Never decoration, no gradients, no glow.
- **Type:** DM Sans for human text, JetBrains Mono for IDs / URLs / timestamps / counts / the wordmark. Numeric values are weight **500**, never bold.
- **Status is a colored dot + text label** (`.tc-status`), never an emoji or bare swatch.
- **Type indicators are dot-prefix + accent-border tags** (`.tc-tag`), NOT a colored left stripe (the stripe was retired). SharePoint type → accent map: LIST→blue, LIBRARY→green, VIEW→amber, PAGE→purple, SYSTEM→red, LINK→cyan.
- Radii 6/8/12/16/pill; 4px spacing grid; motion 180ms on `cubic-bezier(0.2,0,0,1)`, animate color/background/border only (no scale/bounce).

The standalone HTML **inlines** these tokens (the self-contained constraint forbids linking `theme.css`); the snippet defines them on `:host` inside its shadow root. The HTML also `<link>`s DM Sans + JetBrains Mono from Google Fonts with system fallbacks — in a locked-down GCC tenant the CDN may be blocked, in which case it degrades to Segoe UI / Consolas with no breakage.

## Editor note

`SP.code-workspace` only carries a Material-style terminal color theme; it is not a build or task configuration.
