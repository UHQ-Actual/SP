/*
 * SharePoint List Agent — browser reader
 * ---------------------------------------
 * Paste this whole file into the DevTools Console (F12) while you are ON your
 * SharePoint site, signed in through your org's normal SSO. It reads Lists via
 * SharePoint's own REST API using the session cookie the page already holds —
 * no app registration, no Graph token, no extra login.
 *
 *   SPLA.lists()                 -> print every non-hidden list/library
 *   await SPLA.read('My List')   -> return {web, list, columns, items}
 *   await SPLA.export('My List') -> read + download the JSON to disk
 *
 * Drop the downloaded JSON into the MCP server's export folder
 * (config/exports/) and the coding agent can read it locally.
 *
 * Cross-site reads work only within the SAME tenant host you are running on
 * (same-origin cookie auth). It cannot reach another tenant headlessly.
 */
(function () {
  "use strict";

  var USER_TYPES = { User: 1, UserMulti: 1 };
  var LOOKUP_TYPES = { Lookup: 1, LookupMulti: 1 };
  var SYS_KEEP = { Title: 1, Created: 1, Modified: 1, Author: 1, Editor: 1 };
  var TYPE_SKIP = {
    Computed: 1, Attachments: 1, ContentType: 1, Guid: 1,
    Threading: 1, WorkflowStatus: 1, WorkflowEventType: 1, Recurrence: 1
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
  function person(v) {
    if (!v) return "";
    var t = v.Title || "";
    var e = v.EMail || v.Email || "";
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
  function flattenLookup(v) {
    if (v == null) return "";
    if (Array.isArray(v) || (v && v.results)) return multi(v).map(function (e) { return e && e.Title != null ? e.Title : e; }).filter(function (x) { return x !== "" && x != null; }).join("; ");
    return v.Title != null ? v.Title : v;
  }
  function flattenScalar(type, v) {
    if (v == null) return "";
    if (type === "Boolean") return v ? "Yes" : "No";
    if (type === "URL" && typeof v === "object") return v.Description ? v.Description + " (" + v.Url + ")" : v.Url;
    if (type === "MultiChoice") return multi(v).join("; ");
    return v;
  }

  async function fields(w, title) {
    var url = w + "/_api/web/lists/getByTitle('" + escTitle(title) +
      "')/fields?$select=Title,InternalName,TypeAsString,Hidden,CanBeDeleted&$top=500";
    var j = await spFetch(url);
    return (j.value || []).filter(function (f) {
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
    lookups.forEach(function (f) { exp.push(f.InternalName); sel.push(f.InternalName + "/Title", f.InternalName + "/Id"); });

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
      raw = await pageAll(base + "?$select=" + flatSel.join(",") + "&$top=" + PAGE);
    }

    var items = raw.map(function (r) {
      var o = { Id: r.Id };
      scalars.forEach(function (f) { o[f.Title] = flattenScalar(f.TypeAsString, r[f.InternalName]); });
      users.forEach(function (f) { o[f.Title] = expanded ? flattenUser(r[f.InternalName]) : (r[f.InternalName + "Id"] != null ? "id:" + r[f.InternalName + "Id"] : ""); });
      lookups.forEach(function (f) { o[f.Title] = expanded ? flattenLookup(r[f.InternalName]) : (r[f.InternalName + "Id"] != null ? "id:" + r[f.InternalName + "Id"] : ""); });
      return o;
    });

    var columns = [{ name: "ID", internal: "Id", type: "Counter" }].concat(
      fs.map(function (f) { return { name: f.Title, internal: f.InternalName, type: f.TypeAsString }; })
    );

    var out = {
      web: w,
      list: title,
      listId: meta.Id || null,
      exportedAt: new Date().toISOString(),
      columns: columns,
      items: items
    };
    console.log("[SPLA] Read " + items.length + " rows, " + columns.length + " columns" + (expanded ? "" : " (flat fallback)"));
    return out;
  }

  function download(name, obj) {
    var blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  async function SPLA_export(title) {
    var data = await SPLA_read(title);
    var safe = String(title).replace(/[^A-Za-z0-9._-]+/g, "_");
    var day = data.exportedAt.slice(0, 10).replace(/-/g, "");
    var name = "sp_" + safe + "_" + day + ".json";
    download(name, data);
    console.log("[SPLA] Downloaded " + name + " (" + data.items.length + " rows) — move it to the server's config/exports/ folder.");
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

  window.SPLA = { web: web, lists: SPLA_lists, read: SPLA_read, export: SPLA_export };
  console.log("%c[SPLA] ready", "font-weight:bold");
  console.log("  SPLA.lists()                 list non-hidden lists/libraries");
  console.log("  await SPLA.read('My List')   read -> {web,list,columns,items}");
  console.log("  await SPLA.export('My List') read + download JSON to disk");
})();
