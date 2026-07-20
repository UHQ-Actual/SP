/*
 * SharePoint List Agent — browser reader
 * ---------------------------------------
 * Paste this whole file into the DevTools Console (F12) while you are ON your
 * SharePoint site, signed in through your org's normal SSO. It reads Lists via
 * SharePoint's own REST API using the session cookie the page already holds —
 * no app registration, no Graph token, no extra login.
 *
 *   SPLA.lists()                 -> print every non-hidden list/library
 *   await SPLA.read('My List')   -> {web, list, listId, exportedAt, expanded,
 *                                    [degraded], columns, items}
 *   await SPLA.export('My List') -> read + write/download the JSON to disk
 *
 * `expanded:false` means the expanded query failed and person/lookup columns
 * hold "id:N" instead of names — `degraded` spells that out for the reader.
 *
 * Drop the downloaded JSON into the MCP server's export folder
 * (config/exports/) and the coding agent can read it locally.
 *
 * await SPLA.folder() must be typed as the TOP-LEVEL console statement: the
 * folder picker needs transient user activation and will throw SecurityError
 * if it is called from a timer or after some other await.
 *
 * Cross-site reads work only within the SAME tenant host you are running on
 * (same-origin cookie auth). It cannot reach another tenant headlessly.
 */
(function () {
  "use strict";

  var USER_TYPES = { User: 1, UserMulti: 1 };
  var LOOKUP_TYPES = { Lookup: 1, LookupMulti: 1 };
  var SYS_KEEP = { Title: 1, Created: 1, Modified: 1, Author: 1, Editor: 1, FileLeafRef: 1 };
  var TYPE_SKIP = {
    Computed: 1, Attachments: 1, ContentType: 1, Guid: 1,
    Threading: 1, WorkflowStatus: 1, WorkflowEventType: 1, Recurrence: 1,
    // Multi-value managed metadata cannot be fetched via $select on /items
    // (documented SharePoint REST limitation — the query 400s).
    TaxonomyFieldTypeMulti: 1
  };
  var PAGE = 500;
  var MAX_RETRY = 5;

  function web() {
    var ctx = window._spPageContextInfo;
    if (ctx && ctx.webAbsoluteUrl) return ctx.webAbsoluteUrl.replace(/\/$/, "");
    console.warn("[SPLA] No _spPageContextInfo — are you on a SharePoint page? Falling back to origin.");
    return location.origin;
  }

  function escTitle(t) { return encodeURIComponent(String(t).replace(/'/g, "''")); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // Fetch with same-origin cookies, clean nometadata JSON, and 429/503 backoff.
  async function spFetch(url) {
    for (var attempt = 0; ; attempt++) {
      var res = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata" },
        credentials: "same-origin"
      });
      if (res.status === 429 || res.status === 503) {
        if (attempt >= MAX_RETRY) throw new Error("Throttled (" + res.status + "); gave up after " + MAX_RETRY + " retries");
        var wait = parseInt(res.headers.get("Retry-After") || "5", 10);
        console.warn("[SPLA] Throttled " + res.status + "; waiting " + wait + "s (attempt " + (attempt + 1) + ")");
        await sleep(wait * 1000);
        continue;
      }
      if (!res.ok) {
        var body = await res.text();
        throw new Error("HTTP " + res.status + " for " + url + " :: " + body.slice(0, 240));
      }
      return res.json();
    }
  }

  // Follow odata.nextLink to page through a full collection.
  async function pageAll(url) {
    var rows = [];
    while (url) {
      var j = await spFetch(url);
      if (j && j.value) rows = rows.concat(j.value);
      url = j["odata.nextLink"] || j["@odata.nextLink"] || (j.d && j.d.__next) || null;
    }
    return rows;
  }

  // ---- flatteners: turn nested REST shapes into readable strings ----
  // Last line of defence: nothing that reaches `items` may be a raw object.
  function str(v) {
    if (v == null) return "";
    if (typeof v === "object") return JSON.stringify(v);
    return String(v);
  }
  function person(v) {
    if (!v) return "";
    var t = str(v.Title);
    var e = str(v.EMail || v.Email);
    return e && e !== t ? t + " <" + e + ">" : t;
  }
  function multi(v) {
    if (v == null) return [];
    if (Array.isArray(v)) return v;
    if (Array.isArray(v.results)) return v.results;
    return [v];
  }
  function flattenUser(v) {
    if (v == null) return "";
    if (Array.isArray(v) || (v && v.results)) return multi(v).map(person).filter(Boolean).join("; ");
    return person(v);
  }
  function flattenLookup(v, show) {
    if (v == null) return "";
    function pick(e) {
      if (e == null) return "";
      // str(): a lookup can project another nested shape — never let it through raw.
      return e[show] != null ? str(e[show]) : (e.Title != null ? str(e.Title) : "");
    }
    if (Array.isArray(v) || (v && v.results)) return multi(v).map(pick).filter(function (x) { return x !== "" && x != null; }).join("; ");
    return pick(v);
  }
  function flattenScalar(type, v) {
    if (v == null) return "";
    if (type === "Boolean") return v ? "Yes" : "No";
    if (type === "URL" && typeof v === "object") {
      // v.Url can be absent on a half-filled hyperlink field; keep it a string
      // so the column never vanishes from the JSON.
      var u = str(v.Url);
      return v.Description ? str(v.Description) + " (" + u + ")" : u;
    }
    if (type === "MultiChoice") return multi(v).map(str).join("; ");
    if (type === "TaxonomyFieldType") return str(v && v.Label);
    if (type === "Geolocation") return v && v.Latitude != null ? str(v.Latitude) + ", " + str(v.Longitude) : "";
    if (typeof v === "object") return JSON.stringify(v); // never let a raw object land in items
    return v;
  }

  async function fields(w, title) {
    var url = w + "/_api/web/lists/getByTitle('" + escTitle(title) +
      "')/fields?$select=Title,InternalName,TypeAsString,Hidden,CanBeDeleted,LookupField&$top=500";
    var all = await pageAll(url); // page: /fields can exceed 500 on wide lists
    return all.filter(function (f) {
      if (f.Hidden) return false;
      if (String(f.InternalName || "").charAt(0) === "_") return false;
      if (TYPE_SKIP[f.TypeAsString]) return false;
      return f.CanBeDeleted || SYS_KEEP[f.InternalName];
    });
  }

  async function listMeta(w, title) {
    var url = w + "/_api/web/lists/getByTitle('" + escTitle(title) + "')?$select=Id,Title,ItemCount";
    try { return await spFetch(url); } catch (e) { return {}; }
  }

  async function SPLA_read(title) {
    var w = web();
    console.log("[SPLA] Reading '" + title + "' from " + w);
    var fs = await fields(w, title);
    var meta = await listMeta(w, title);

    var scalars = [], users = [], lookups = [];
    fs.forEach(function (f) {
      if (USER_TYPES[f.TypeAsString]) users.push(f);
      else if (LOOKUP_TYPES[f.TypeAsString]) lookups.push(f);
      else scalars.push(f);
    });

    var sel = ["Id"], exp = [];
    scalars.forEach(function (f) { sel.push(f.InternalName); });
    users.forEach(function (f) { exp.push(f.InternalName); sel.push(f.InternalName + "/Title", f.InternalName + "/EMail"); });
    lookups.forEach(function (f) {
      // Project the column the lookup actually shows (ShowField), not always Title.
      f.showField = f.LookupField || "Title";
      exp.push(f.InternalName);
      sel.push(f.InternalName + "/" + f.showField, f.InternalName + "/Id");
    });

    var base = w + "/_api/web/lists/getByTitle('" + escTitle(title) + "')/items";
    var full = base + "?$select=" + sel.join(",") + (exp.length ? "&$expand=" + exp.join(",") : "") + "&$top=" + PAGE;

    var raw, expanded = true;
    try {
      raw = await pageAll(full);
    } catch (e) {
      console.warn("[SPLA] Expanded query failed (" + e.message + "); retrying flat.");
      expanded = false;
      var flatSel = ["Id"].concat(scalars.map(function (f) { return f.InternalName; }))
        .concat(users.map(function (f) { return f.InternalName + "Id"; }))
        .concat(lookups.map(function (f) { return f.InternalName + "Id"; }));
      try {
        raw = await pageAll(base + "?$select=" + flatSel.join(",") + "&$top=" + PAGE);
      } catch (e2) {
        throw new Error("Read of '" + title + "' failed — expanded query: " + e.message + " | flat fallback: " + e2.message);
      }
    }

    var items = raw.map(function (r) {
      var o = { Id: r.Id };
      scalars.forEach(function (f) { o[f.Title] = flattenScalar(f.TypeAsString, r[f.InternalName]); });
      users.forEach(function (f) { o[f.Title] = expanded ? flattenUser(r[f.InternalName]) : (r[f.InternalName + "Id"] != null ? "id:" + r[f.InternalName + "Id"] : ""); });
      lookups.forEach(function (f) { o[f.Title] = expanded ? flattenLookup(r[f.InternalName], f.showField) : (r[f.InternalName + "Id"] != null ? "id:" + r[f.InternalName + "Id"] : ""); });
      return o;
    });

    var columns = [{ name: "ID", internal: "Id", type: "Counter" }].concat(
      fs.map(function (f) { return { name: f.Title, internal: f.InternalName, type: f.TypeAsString }; })
    );

    // Record fidelity IN the payload: downstream (the MCP server, the agent)
    // otherwise cannot tell a real value from the "id:12" a flat fallback
    // writes into every person and lookup column.
    var out = {
      web: w,
      list: title,
      listId: meta.Id || null,
      exportedAt: new Date().toISOString(),
      expanded: expanded
    };
    if (!expanded) out.degraded = "person/lookup columns are ids only";
    out.columns = columns;
    out.items = items;
    console.log("[SPLA] Read " + items.length + " rows, " + columns.length + " columns" + (expanded ? "" : " (flat fallback)"));
    return out;
  }

  // ---- output: chosen folder (File System Access) or plain download ----
  // Point this at the server's config/exports/ once and exports land there
  // directly — no download, no moving files. Edge/Chrome only; the grant lasts
  // one page-load (browser rule). A write failure is retried once and then falls
  // back to download for that export alone — the grant survives transient locks.
  var outDir = null;
  var recheckPerm = false;   // set after a failed write; the next export re-verifies the grant
  var RETRY_MS = 400;

  async function SPLA_folder() {
    if (!window.showDirectoryPicker) {
      console.warn("[SPLA] No File System Access API in this browser — exports will download normally.");
      return false;
    }
    // A cancelled or SecurityError'd picker must NOT cost the user a grant they
    // already have — assign only on success, restore the previous handle otherwise.
    var prev = outDir;
    try {
      var picked = await window.showDirectoryPicker({ mode: "readwrite" });
      outDir = picked;
      recheckPerm = false;
      console.log("[SPLA] Export folder set — exports write straight into it (this page-load only).");
      return true;
    } catch (e) {
      outDir = prev;
      if (prev) {
        console.warn("[SPLA] Folder picker cancelled/failed (" + e.message + ") — KEEPING the folder already granted; exports still write there.");
        return true;
      }
      console.warn("[SPLA] Folder not set: " + e.message + " — exports will download. (Run `await SPLA.folder()` as the top-level console statement; it needs user activation.)");
      return false;
    }
  }

  function download(name, text) {
    var blob = new Blob([text], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  // Re-confirm a grant that a previous write failed under, before touching disk
  // again. queryPermission/requestPermission are optional on older engines.
  async function permissionOk() {
    if (!outDir || !recheckPerm) return true;
    recheckPerm = false;
    if (typeof outDir.queryPermission !== "function") return true;
    try {
      var st = await outDir.queryPermission({ mode: "readwrite" });
      if (st !== "granted" && typeof outDir.requestPermission === "function") {
        st = await outDir.requestPermission({ mode: "readwrite" });
      }
      if (st !== "granted") {
        outDir = null;
        console.error("[SPLA] Export folder permission is now '" + st + "' — the grant is gone. Re-run `await SPLA.folder()`; this export downloads instead.");
        return false;
      }
      console.log("[SPLA] Export folder permission re-confirmed after the earlier write failure.");
      return true;
    } catch (e) {
      console.warn("[SPLA] Could not re-check folder permission (" + e.message + "); attempting the write anyway.");
      return true;
    }
  }

  // One folder-write attempt. Resolves null on success, or the Error that
  // stopped it. It NEVER leaves a partial or 0-byte file behind: getFileHandle
  // (create:true) creates the entry before any bytes are written, so on failure
  // the stream is aborted (releasing Chrome's .crswap lock) and the entry we
  // ourselves created is removed. A half-written export the MCP server can't
  // parse is worse than no export at all — it makes store.find() serve
  // yesterday's rows silently.
  //
  // WHY THE PRE-EXISTENCE PROBE — do not "simplify" it away:
  // cleanup must delete ONLY an entry this call brought into existence.
  //   * fresh create + failure -> the 0-byte entry is ours, remove it.
  //   * OVERWRITE + failure    -> the entry is the user's PREVIOUS good export.
  //     ws.abort() already discards the swap file and leaves that export
  //     byte-for-byte intact, so removeEntry() here would destroy a good file
  //     over a transient lock (VS Code Live Preview, OneDrive, AV scan) — far
  //     worse than the 0-byte leftover the cleanup was added to fix.
  // The probe is a plain getFileHandle() with no create flag: it throws
  // NotFoundError when the name is free, which is exactly the signal we want.
  async function attemptFolderWrite(name, text) {
    var ws = null;
    var preExisting = false;
    try { await outDir.getFileHandle(name); preExisting = true; } catch (e) { preExisting = false; }
    try {
      var fh = await outDir.getFileHandle(name, { create: true });
      ws = await fh.createWritable();
      await ws.write(text);
      await ws.close();
      return null;
    } catch (e) {
      if (ws && typeof ws.abort === "function") {
        try { await ws.abort(); } catch (_) { /* already closed/aborted */ }
      }
      if (!preExisting) {
        try {
          if (outDir && typeof outDir.removeEntry === "function") await outDir.removeEntry(name);
        } catch (_) { /* nothing was created, or the engine lacks removeEntry */ }
      }
      return e;
    }
  }

  // Single choke point for every export: chosen folder first, download as fallback.
  async function writeOut(name, obj) {
    // Serialize ONCE, before the filesystem is touched. If this throws it is the
    // payload's fault (RangeError on a huge list), not the folder's — say so and
    // keep the grant, instead of blaming the folder and then throwing again from
    // the download path.
    var text;
    try {
      text = JSON.stringify(obj, null, 2);
    } catch (e) {
      throw new Error("Export of '" + ((obj && obj.list) || "list") +
        "' could not be serialized (" + e.message + "); nothing was written and the export folder is unchanged.");
    }

    if (outDir && await permissionOk()) {
      var err = await attemptFolderWrite(name, text);
      if (!err) return "folder";
      // Transient locks are the common case here (VS Code Live Preview, OneDrive
      // sync, AV scan). Retry once before giving up on this export.
      console.warn("[SPLA] Folder write failed (" + err.message + "); retrying once in " + RETRY_MS + "ms.");
      await sleep(RETRY_MS);
      var err2 = await attemptFolderWrite(name, text);
      if (!err2) {
        console.log("[SPLA] Folder write succeeded on retry — export landed in the chosen folder.");
        return "folder";
      }
      // Fall back for THIS export only. The handle is kept so one bad moment
      // does not silently downgrade the whole page-load to Downloads.
      recheckPerm = true;
      console.error("[SPLA] Folder write failed twice (" + err2.message + "); falling back to download for THIS export. " +
        "The folder is KEPT — the next export re-checks the permission first. If it keeps failing, re-run `await SPLA.folder()`.");
    }
    download(name, text);
    return "download";
  }

  // Filename rule — injective enough that two lists never overwrite each other.
  // Keep the readable sanitized stem, and append a 6-hex hash of the RAW title
  // whenever sanitizing changed or emptied it, because that is exactly when two
  // different titles can collapse onto one stem: 'Cases (2026)', 'Cases 2026'
  // and 'Cases_2026' all sanitize to 'Cases_2026', and '日本語'/'中文' both
  // sanitize to nothing. A title that survives sanitizing untouched is already
  // unique, so it keeps the clean name. Same list + same day still overwrites on
  // purpose: a re-export is a refresh of one artifact, not a new one.
  function hash6(s) {
    var h = 0x811c9dc5;                                   // FNV-1a, 32-bit
    for (var i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("00000000" + h.toString(16)).slice(-6);
  }

  function exportName(title, iso) {
    var raw = String(title);
    var safe = raw.replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "");
    var lossy = safe !== raw;
    if (!safe) safe = "list";
    var day = iso.slice(0, 10).replace(/-/g, "");
    return "sp_" + safe + "_" + day + (lossy ? "_" + hash6(raw) : "") + ".json";
  }

  async function SPLA_export(title) {
    var data = await SPLA_read(title);
    var name = exportName(title, data.exportedAt);
    var how = await writeOut(name, data);
    console.log("[SPLA] " + (how === "folder"
      ? "Wrote " + name + " into the chosen folder"
      : "Downloaded " + name + " — move it to the server's config/exports/ folder")
      + " (" + data.items.length + " rows).");
    return data;
  }

  async function SPLA_lists() {
    var w = web();
    var url = w + "/_api/web/lists?$select=Title,ItemCount,BaseTemplate,Hidden&$filter=Hidden eq false&$top=500";
    var j = await spFetch(url);
    var rows = (j.value || []).map(function (l) { return { Title: l.Title, Items: l.ItemCount, Template: l.BaseTemplate }; });
    rows.sort(function (a, b) { return String(a.Title).localeCompare(String(b.Title)); });
    console.table(rows);
    console.log("[SPLA] " + rows.length + " lists. Use await SPLA.export('<Title>') to pull one.");
    return rows;
  }

  window.SPLA = { web: web, lists: SPLA_lists, read: SPLA_read, export: SPLA_export, folder: SPLA_folder };
  console.log("%c[SPLA] ready", "font-weight:bold");
  console.log("  await SPLA.folder()          pick config\\exports\\ once — exports write straight there");
  console.log("  SPLA.lists()                 list non-hidden lists/libraries");
  console.log("  await SPLA.read('My List')   read -> {web,list,columns,items}");
  console.log("  await SPLA.export('My List') read + write/download the JSON");
})();
