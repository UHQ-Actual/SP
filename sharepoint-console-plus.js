/*
 * SharePoint Console Plus
 * -----------------------
 * One in-page panel that merges the standalone link launcher (Links tab) with a
 * live REST toolkit for the current SharePoint list (Overview / Columns / Views /
 * Recycle Bin / Export / Versions).
 *
 * HOW TO RUN
 *   Open a SharePoint list page, then paste this file into the DevTools Console,
 *   or save it as a DevTools Snippet (Sources > Snippets) and Run. Re-running
 *   toggles the panel off/on (guarded via window.__SPCP).
 *
 * WHY IN-PAGE
 *   The REST tabs call /_api/... using your session cookies, which browsers only
 *   allow same-origin, inside the SharePoint page. So this cannot ship as a
 *   standalone .html file. On a non-SharePoint page the panel still opens but the
 *   REST tabs disable themselves; only Links stays active.
 *
 * STORAGE
 *   The Links tab shares localStorage key `sp_link_console_v1` with the standalone
 *   sharepoint-console.html, so exports move cleanly between the two.
 */
(function () {
  "use strict";

  /* ================= double-injection guard ================= */
  if (window.__SPCP && typeof window.__SPCP.toggle === "function") {
    window.__SPCP.toggle();
    return;
  }

  var KEY = "sp_link_console_v1";
  var TYPES = ["LIST", "LIBRARY", "VIEW", "PAGE", "SYSTEM", "LINK"];
  // Default site to crawl for lists (WHD · Minimum Wage planning).
  var DEFAULT_BASE = "https://usdol.sharepoint.com/sites/whd/mw/mwplanning/";

  /* ================= context detection ================= */
  var ctx = detectContext();
  function detectContext() {
    var c = { webUrl: "", webServerRel: "", listId: "", listTitle: "", restOk: false };
    var info = window._spPageContextInfo;
    if (info && info.webAbsoluteUrl) {
      c.webUrl = info.webAbsoluteUrl.replace(/\/$/, "");
      c.webServerRel = (info.webServerRelativeUrl || "").replace(/\/$/, "");
      c.listId = info.pageListId ? String(info.pageListId).replace(/[{}]/g, "") : "";
      c.restOk = true;
    } else {
      // Fallback: parse /sites/ or /teams/ from the address.
      var m = location.pathname.match(/^(.*?\/(?:sites|teams)\/[^\/]+)/i);
      if (m) {
        c.webServerRel = m[1].replace(/\/$/, "");
        c.webUrl = location.origin + c.webServerRel;
        c.restOk = /\.sharepoint\./i.test(location.hostname);
      }
    }
    return c;
  }

  /* ================= REST helpers ================= */
  function apiHeaders(extra) {
    var h = { "Accept": "application/json;odata=nometadata" };
    if (extra) for (var k in extra) h[k] = extra[k];
    return h;
  }
  function spGet(rel) {
    var url = rel.indexOf("http") === 0 ? rel : ctx.webUrl + rel;
    return fetch(url, { headers: apiHeaders(), credentials: "same-origin" })
      .then(function (r) {
        if (!r.ok) throw new Error("GET " + r.status + " " + rel);
        return r.json();
      });
  }
  function spGetAll(rel) {
    var acc = [];
    function page(u) {
      return spGet(u).then(function (data) {
        acc = acc.concat(data.value || []);
        var next = data["odata.nextLink"] || data["@odata.nextLink"];
        return next ? page(next) : acc;
      });
    }
    return page(rel);
  }
  function getDigest() {
    return fetch(ctx.webUrl + "/_api/contextinfo", {
      method: "POST",
      headers: apiHeaders(),
      credentials: "same-origin"
    }).then(function (r) { return r.json(); })
      .then(function (d) { return d.FormDigestValue || (d.d && d.d.GetContextWebInformation.FormDigestValue); });
  }
  function spPost(rel, body, opts) {
    opts = opts || {};
    return getDigest().then(function (digest) {
      var headers = apiHeaders({
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest
      });
      if (opts.merge) {
        headers["X-HTTP-Method"] = "MERGE";
        headers["IF-MATCH"] = "*";
      }
      return fetch(ctx.webUrl + rel, {
        method: "POST",
        headers: headers,
        credentials: "same-origin",
        body: body ? JSON.stringify(body) : undefined
      });
    }).then(function (r) {
      if (!r.ok) throw new Error("POST " + r.status + " " + rel);
      return r.status === 204 ? {} : r.json().catch(function () { return {}; });
    });
  }
  function listApi() {
    return ctx.listId
      ? "/_api/web/lists(guid'" + ctx.listId + "')"
      : "/_api/web/lists/getbytitle('" + encodeURIComponent(ctx.listTitle) + "')";
  }

  /* ================= launcher storage + parser ================= */
  var links = [];
  var storageOk = true;
  function loadLinks() {
    try { links = JSON.parse(localStorage.getItem(KEY)) || []; }
    catch (e) { storageOk = false; links = []; }
  }
  function saveLinks() {
    try { localStorage.setItem(KEY, JSON.stringify(links)); }
    catch (e) { storageOk = false; }
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function pretty(s) {
    try { s = decodeURIComponent(s); } catch (e) {}
    return s.replace(/[-_]+/g, " ").replace(/\.aspx$/i, "").trim();
  }
  function titleCase(s) {
    return s.replace(/\w\S*/g, function (w) { return w.charAt(0).toUpperCase() + w.slice(1); });
  }
  function rid() { return Math.random().toString(36).slice(2, 10); }
  function lastAspx(parts) {
    for (var i = parts.length - 1; i >= 0; i--) if (/\.aspx$/i.test(parts[i])) return pretty(parts[i]);
    return "";
  }
  function parse(url) {
    var out = { url: url, site: "", name: "", type: "LINK", view: "" }, u;
    try { u = new URL(url); } catch (e) { out.site = "Unparsed"; out.name = url; return out; }
    var parts = u.pathname.split("/").filter(Boolean);
    var lower = parts.map(function (p) { return p.toLowerCase(); });
    var si = lower.indexOf("sites"); if (si < 0) si = lower.indexOf("teams");
    out.site = (si >= 0 && parts[si + 1]) ? titleCase(pretty(parts[si + 1])) : u.hostname;
    var li = lower.indexOf("lists"), fi = lower.indexOf("forms");
    if (lower.indexOf("_layouts") >= 0) {
      out.type = "SYSTEM"; out.name = lastAspx(parts) || "System page";
    } else if (li >= 0 && parts[li + 1]) {
      var listName = titleCase(pretty(parts[li + 1])), file = lower[li + 2] || "";
      if (file && file.indexOf("allitems") < 0 && /\.aspx$/.test(file)) {
        out.type = "VIEW"; out.view = pretty(parts[li + 2]); out.name = listName + " · " + titleCase(out.view);
      } else { out.type = "LIST"; out.name = listName; }
    } else if (fi >= 0 && parts[fi - 1]) {
      out.type = "LIBRARY"; out.name = titleCase(pretty(parts[fi - 1]));
    } else if (lower.indexOf("sitepages") >= 0) {
      out.type = "PAGE"; out.name = titleCase(lastAspx(parts) || "Page");
    } else if (lastAspx(parts)) {
      out.type = "PAGE"; out.name = titleCase(lastAspx(parts));
    } else {
      out.type = "LINK";
      out.name = parts.length ? titleCase(pretty(parts[parts.length - 1])) : u.hostname;
    }
    if (!out.name) out.name = u.hostname;
    return out;
  }
  function makeItem(url) { var p = parse(url); p.id = rid(); p.label = ""; return p; }
  function displayName(it) { return it.label || it.name; }
  function addUrls(text) {
    var found = (text.match(/https?:\/\/[^\s"'<>]+/gi) || [])
      .map(function (s) { return s.replace(/[.,);]+$/, ""); });
    if (!found.length) { log("No URLs found."); return; }
    var seen = {}; links.forEach(function (it) { seen[it.url] = 1; });
    var added = 0;
    found.forEach(function (url) { if (!seen[url]) { seen[url] = 1; links.push(makeItem(url)); added++; } });
    saveLinks(); renderLinks();
    log(added ? "Added " + added + " link(s)" : "Already present.");
  }

  /* ================= UI shell (Shadow DOM) ================= */
  var host = document.createElement("div");
  host.id = "spcp-host";
  host.style.cssText = "position:fixed;top:70px;right:24px;z-index:2147483000;";
  var root = host.attachShadow({ mode: "open" });
  document.documentElement.appendChild(host);

  var style = document.createElement("style");
  style.textContent = cssText();
  root.appendChild(style);

  var panel = el("div", "panel");
  root.appendChild(panel);

  var header = el("div", "hdr");
  header.innerHTML = '<span class="title">sharepoint.console<b>.plus</b></span>' +
    '<span class="ctx"></span><span class="spacer"></span>' +
    '<button class="x" title="Close">×</button>';
  panel.appendChild(header);
  header.querySelector(".ctx").textContent = ctx.restOk
    ? (ctx.listId ? "list bound" : "web only") : "no SharePoint context";
  header.querySelector(".x").onclick = function () { api.toggle(); };

  var tabsBar = el("div", "tabs");
  panel.appendChild(tabsBar);

  var bodyEl = el("div", "body");
  panel.appendChild(bodyEl);

  var logEl = el("div", "log");
  panel.appendChild(logEl);

  var TABS = [
    { id: "links",   label: "Links",      rest: false, render: tabLinks },
    { id: "gulp",    label: "Gulp",       rest: true,  render: tabGulp },
    { id: "overview",label: "Overview",   rest: true,  render: tabOverview },
    { id: "columns", label: "Columns",    rest: true,  render: tabColumns },
    { id: "views",   label: "Views",      rest: true,  render: tabViews },
    { id: "recycle", label: "Recycle Bin",rest: true,  render: tabRecycle },
    { id: "export",  label: "Export",     rest: true,  render: tabExport },
    { id: "versions",label: "Versions",   rest: true,  render: tabVersions }
  ];
  var active = "links";
  TABS.forEach(function (t) {
    var b = el("button", "tab");
    b.textContent = t.label;
    b.dataset.id = t.id;
    var disabled = t.rest && !ctx.restOk;
    if (disabled) { b.classList.add("disabled"); b.title = "Needs a SharePoint page"; }
    b.onclick = function () { if (!disabled) selectTab(t.id); };
    tabsBar.appendChild(b);
  });

  function selectTab(id) {
    active = id;
    Array.prototype.forEach.call(tabsBar.children, function (b) {
      b.classList.toggle("active", b.dataset.id === id);
    });
    bodyEl.innerHTML = "";
    var t = TABS.filter(function (x) { return x.id === id; })[0];
    try { t.render(bodyEl); }
    catch (e) { bodyEl.appendChild(note("Error: " + e.message)); }
  }

  /* dragging */
  (function () {
    var dragging = false, ox = 0, oy = 0;
    header.addEventListener("mousedown", function (e) {
      if (e.target.tagName === "BUTTON") return;
      dragging = true;
      var r = host.getBoundingClientRect();
      ox = e.clientX - r.left; oy = e.clientY - r.top;
      e.preventDefault();
    });
    window.addEventListener("mousemove", function (e) {
      if (!dragging) return;
      host.style.left = (e.clientX - ox) + "px";
      host.style.top = (e.clientY - oy) + "px";
      host.style.right = "auto";
    });
    window.addEventListener("mouseup", function () { dragging = false; });
  })();

  /* ================= tab: Links ================= */
  function tabLinks(container) {
    var bar = el("div", "row");
    var ta = el("textarea"); ta.placeholder = "Paste SharePoint URLs…";
    var addBtn = btn("Add", function () { addUrls(ta.value); ta.value = ""; });
    var addThis = btn("Add this list", function () {
      var url = location.href.split("#")[0];
      addUrls(url);
    });
    var exp = btn("Export", exportLinks);
    var imp = btn("Import", importLinks);
    bar.appendChild(addBtn); bar.appendChild(addThis); bar.appendChild(exp); bar.appendChild(imp);
    var search = el("input"); search.placeholder = "Search…"; search.className = "search";
    search.oninput = function () { renderLinks(search.value); };

    container.appendChild(ta);
    container.appendChild(bar);
    container.appendChild(search);
    var mount = el("div", "links-mount");
    container.appendChild(mount);
    linksMount = mount;
    renderLinks(search.value);
  }
  var linksMount = null, linkFilter = "";
  function renderLinks(filter) {
    if (typeof filter === "string") linkFilter = filter;
    if (!linksMount) return;
    linksMount.innerHTML = "";
    var q = (linkFilter || "").toLowerCase();
    var rows = links.filter(function (it) {
      return !q || (displayName(it) + " " + it.site + " " + it.url).toLowerCase().indexOf(q) >= 0;
    });
    if (!links.length) { linksMount.appendChild(note("No saved links. Paste some above or use “Add this list”.")); return; }
    if (!rows.length) { linksMount.appendChild(note("No links match.")); return; }
    var groups = {};
    rows.forEach(function (it) { (groups[it.site] = groups[it.site] || []).push(it); });
    Object.keys(groups).sort().forEach(function (site) {
      var h = el("div", "grp-h"); h.textContent = site + " (" + groups[site].length + ")";
      linksMount.appendChild(h);
      groups[site].forEach(function (it) { linksMount.appendChild(linkCard(it)); });
    });
  }
  function linkCard(it) {
    var c = el("div", "lcard"); c.dataset.t = it.type;
    var n = el("div", "lname"); n.textContent = displayName(it); n.title = "Double-click to rename";
    n.ondblclick = function () { n.contentEditable = "true"; n.focus(); };
    n.onblur = function () {
      n.contentEditable = "false";
      var v = n.textContent.trim();
      it.label = (v && v !== it.name) ? v : ""; n.textContent = displayName(it); saveLinks();
    };
    n.onkeydown = function (e) { if (e.key === "Enter") { e.preventDefault(); n.blur(); } };
    var meta = el("div", "lmeta"); meta.innerHTML = '<span class="tag">' + esc(it.type) + '</span>';
    var acts = el("div", "lacts");
    acts.appendChild(btn("Open", function () { window.open(it.url, "_blank", "noopener"); }));
    acts.appendChild(btn("Copy", function () { copyText(it.url); }));
    acts.appendChild(btn("×", function () {
      links = links.filter(function (x) { return x.id !== it.id; }); saveLinks(); renderLinks();
    }));
    c.appendChild(n); c.appendChild(meta); c.appendChild(acts);
    return c;
  }
  function exportLinks() {
    downloadBlob(JSON.stringify(links, null, 2), "sharepoint-console-links.json", "application/json");
    log("Exported " + links.length + " links");
  }
  function importLinks() {
    pickFile(function (text) {
      try {
        var arr = JSON.parse(text);
        if (!Array.isArray(arr)) throw 0;
        var seen = {}; links.forEach(function (it) { seen[it.url] = 1; });
        var added = 0;
        arr.forEach(function (it) {
          if (!it || !it.url || seen[it.url]) return; seen[it.url] = 1;
          if (!it.type || !it.id) { var p = makeItem(it.url); p.label = it.label || ""; it = p; }
          links.push(it); added++;
        });
        saveLinks(); renderLinks(); log("Imported " + added + " new link(s)");
      } catch (e) { log("Import failed — invalid JSON"); }
    });
  }

  /* ================= tab: Gulp =================
     Crawl a site for its lists (and any list for its items), render the result
     as a table, and save it for VS Code: Markdown (native preview), JSON, HTML.
     Discovered list URLs are also saved into the Links catalog / JSON store. */
  var lastGulp = null; // { title, cols, rows, kind }
  function tabGulp(container) {
    var wrap = el("div", "col"); wrap.style.gap = "10px";

    container.appendChild(note("Paste a site URL to pull every list + library in it. Runs on your login, same tenant only."));

    var base = el("input"); base.className = "search"; base.value = DEFAULT_BASE;
    base.placeholder = "https://usdol.sharepoint.com/sites/whd/mw/mwplanning/";
    container.appendChild(base);

    var bar = el("div", "row");
    bar.appendChild(btn("Gulp lists", function () { gulpLists(baseOf(base.value)); }));
    var save = el("div", "row"); save.style.marginLeft = "auto";
    var mdBtn = btn("Save .md", function () { saveGulp("md"); });
    var jsonBtn = btn("Save .json", function () { saveGulp("json"); });
    var htmlBtn = btn("Save .html", function () { saveGulp("html"); });
    var csvBtn = btn("Save .csv", function () { saveGulp("csv"); });
    [mdBtn, jsonBtn, htmlBtn, csvBtn].forEach(function (b) { b.classList.add("is-disabled"); save.appendChild(b); });
    bar.appendChild(save);
    container.appendChild(bar);

    var out = el("div"); container.appendChild(out);

    function enableSaves(on) {
      [mdBtn, jsonBtn, htmlBtn, csvBtn].forEach(function (b) { b.classList.toggle("is-disabled", !on); });
    }

    function gulpLists(baseUrl) {
      out.innerHTML = ""; out.appendChild(note("Reading lists from " + baseUrl + " …"));
      enableSaves(false);
      spGetAll(baseUrl + "/_api/web/lists?$select=Title,Id,BaseType,BaseTemplate,ItemCount,Hidden,DefaultViewUrl,LastItemModifiedDate&$top=500")
        .then(function (lists) {
          // Keep visible content lists/libraries; drop hidden system lists and catalog templates (>=116, e.g. app/theme catalogs).
          lists = lists.filter(function (l) { return !l.Hidden && l.Title && !(l.BaseTemplate >= 116); });
          lists.sort(function (a, b) { return a.Title.localeCompare(b.Title); });
          var origin = originOf(baseUrl);
          var savedCount = 0;
          var rows = lists.map(function (l) {
            var url = l.DefaultViewUrl ? origin + l.DefaultViewUrl : baseUrl;
            if (addOne(url)) savedCount++;
            return {
              Title: l.Title,
              Type: l.BaseType === 1 ? "Library" : "List",
              Items: l.ItemCount,
              Modified: (l.LastItemModifiedDate || "").replace("T", " ").replace("Z", ""),
              Url: url,
              _id: l.Id
            };
          });
          if (savedCount) { saveLinks(); }
          lastGulp = { title: "Lists in " + siteLabel(baseUrl), cols: ["Title", "Type", "Items", "Modified", "Url"], rows: rows, kind: "lists" };
          renderTable(out, lastGulp, function (row) { gulpItems(baseUrl, row._id, row.Title); });
          enableSaves(rows.length > 0);
          log("Gulped " + rows.length + " lists · saved " + savedCount + " new URL(s) to Links");
        })
        .catch(function (e) {
          out.innerHTML = "";
          out.appendChild(note("Could not read lists: " + e.message + ". Run this on a page under the same tenant host (e.g. usdol.sharepoint.com)."));
        });
    }

    function gulpItems(baseUrl, listId, title) {
      out.innerHTML = ""; out.appendChild(note("Reading items from “" + title + "” …"));
      enableSaves(false);
      spGetAll(baseUrl + "/_api/web/lists(guid'" + listId + "')/items?$top=2000")
        .then(function (items) {
          var cols = items.length
            ? Object.keys(items[0]).filter(function (k) { return k.indexOf("odata") < 0 && k.indexOf("@") < 0; })
            : [];
          var display = cols.slice(0, 14); // keep the table readable; JSON/CSV keep everything
          var rows = items.map(function (it) {
            var r = {}; cols.forEach(function (c) { r[c] = flatten(it[c]); }); return r;
          });
          lastGulp = { title: title + " — items", cols: display, allCols: cols, rows: rows, kind: "items" };
          renderTable(out, { title: lastGulp.title, cols: display, rows: rows });
          enableSaves(rows.length > 0);
          var extra = cols.length > display.length ? " (showing " + display.length + "/" + cols.length + " columns; full set in JSON/CSV)" : "";
          log("Gulped " + rows.length + " items" + extra);
        })
        .catch(function (e) { out.innerHTML = ""; out.appendChild(note("Could not read items: " + e.message)); });
    }

    function saveGulp(fmt) {
      if (!lastGulp) { log("Nothing gulped yet."); return; }
      var cols = lastGulp.allCols || lastGulp.cols;
      var slug = lastGulp.title.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
      if (fmt === "md") downloadBlob(mdTable(lastGulp.title, cols, lastGulp.rows), slug + ".md", "text/markdown");
      else if (fmt === "json") downloadBlob(JSON.stringify(lastGulp.rows, null, 2), slug + ".json", "application/json");
      else if (fmt === "csv") downloadBlob(toCsv(lastGulp.rows, cols), slug + ".csv", "text/csv");
      else if (fmt === "html") downloadBlob(htmlTable(lastGulp.title, cols, lastGulp.rows), slug + ".html", "text/html");
      log("Saved " + slug + "." + fmt);
    }
  }

  /* base / origin helpers for gulp */
  function baseOf(u) { return String(u || "").trim().replace(/\/+$/, ""); }
  function originOf(u) { try { return new URL(u).origin; } catch (e) { return location.origin; } }
  function siteLabel(u) {
    var m = String(u).match(/\/(?:sites|teams)\/(.+)$/i);
    return m ? m[1].replace(/\/+$/, "") : originOf(u).replace(/^https?:\/\//, "");
  }
  function flatten(v) {
    if (v == null) return "";
    if (typeof v === "object") return v.Title || v.Label || v.Url || v.Description || JSON.stringify(v);
    return v;
  }
  /* add one URL to the Links store without a toast; returns true if newly added */
  function addOne(url) {
    for (var i = 0; i < links.length; i++) if (links[i].url === url) return false;
    links.push(makeItem(url)); return true;
  }

  /* render an interactive table (click-to-sort); optional row-click callback */
  function renderTable(mount, data, onRow) {
    mount.innerHTML = "";
    if (!data.rows.length) { mount.appendChild(note("Empty.")); return; }
    var head = el("div", "grp-h"); head.textContent = data.title + " · " + data.rows.length + " rows";
    mount.appendChild(head);
    var scroller = el("div", "tbl-scroll");
    var tbl = document.createElement("table"); tbl.className = "tbl";
    var thead = document.createElement("thead"); var htr = document.createElement("tr");
    var sortState = { col: null, dir: 1 };
    data.cols.forEach(function (c) {
      var th = document.createElement("th"); th.textContent = c;
      th.onclick = function () {
        sortState.dir = sortState.col === c ? -sortState.dir : 1; sortState.col = c;
        data.rows.sort(function (a, b) {
          var x = a[c], y = b[c];
          var nx = parseFloat(x), ny = parseFloat(y);
          if (!isNaN(nx) && !isNaN(ny)) return (nx - ny) * sortState.dir;
          return String(x).localeCompare(String(y)) * sortState.dir;
        });
        body();
      };
      htr.appendChild(th);
    });
    thead.appendChild(htr); tbl.appendChild(thead);
    var tb = document.createElement("tbody"); tbl.appendChild(tb);
    function body() {
      tb.innerHTML = "";
      data.rows.forEach(function (row) {
        var tr = document.createElement("tr");
        if (onRow) { tr.className = "clickable"; tr.onclick = function () { onRow(row); }; }
        data.cols.forEach(function (c) {
          var td = document.createElement("td");
          var v = row[c] == null ? "" : String(row[c]);
          if (/^https?:\/\//.test(v)) { var a = document.createElement("a"); a.href = v; a.target = "_blank"; a.textContent = v; a.className = "tlink"; td.appendChild(a); }
          else td.textContent = v;
          tr.appendChild(td);
        });
        tb.appendChild(tr);
      });
    }
    body();
    scroller.appendChild(tbl); mount.appendChild(scroller);
    if (onRow) mount.appendChild(note("Tip: click a row to gulp that list's items."));
  }

  /* ---- table serializers (VS Code targets) ---- */
  function mdCell(v) { return String(v == null ? "" : v).replace(/\|/g, "\\|").replace(/\r?\n/g, " "); }
  function mdTable(title, cols, rows) {
    var out = ["# " + title, "", "> " + rows.length + " rows · generated by SharePoint Console Plus", ""];
    out.push("| " + cols.join(" | ") + " |");
    out.push("| " + cols.map(function () { return "---"; }).join(" | ") + " |");
    rows.forEach(function (r) { out.push("| " + cols.map(function (c) { return mdCell(flatten(r[c])); }).join(" | ") + " |"); });
    return out.join("\n") + "\n";
  }
  function htmlTable(title, cols, rows) {
    var esc2 = function (s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); };
    var head = cols.map(function (c) { return "<th>" + esc2(c) + "</th>"; }).join("");
    var bodyRows = rows.map(function (r) {
      return "<tr>" + cols.map(function (c) {
        var v = flatten(r[c]);
        var cell = /^https?:\/\//.test(v) ? '<a href="' + esc2(v) + '" target="_blank">' + esc2(v) + "</a>" : esc2(v);
        return "<td>" + cell + "</td>";
      }).join("") + "</tr>";
    }).join("\n");
    return '<!doctype html><meta charset="utf-8"><title>' + esc2(title) + '</title>' +
      "<style>" +
      "body{background:#0f0f0f;color:#e7ecf3;font:14px 'DM Sans',system-ui,'Segoe UI',sans-serif;margin:0;padding:24px}" +
      "h1{font-size:22px;font-weight:700;letter-spacing:-.014em;margin:0 0 4px}" +
      ".sub{color:#808080;font-size:13px;margin:0 0 16px}" +
      "input{background:#202020;color:#fff;border:1px solid #303030;border-radius:8px;padding:8px 11px;font:inherit;width:280px;margin-bottom:12px}" +
      "input:focus{outline:none;border-color:#5B9FEF}" +
      "table{border-collapse:collapse;width:100%;font-size:13px}" +
      "th,td{text-align:left;padding:7px 10px;border-bottom:1px solid #232323;vertical-align:top}" +
      "th{position:sticky;top:0;background:#141414;color:#a0a0a0;font-size:11px;text-transform:uppercase;letter-spacing:1px;cursor:pointer;user-select:none}" +
      "tr:hover td{background:#161616}" +
      "a{color:#5B9FEF;text-decoration:none;font-family:'JetBrains Mono',Consolas,monospace;font-size:12px}" +
      "a:hover{text-decoration:underline}" +
      "td{font-family:'JetBrains Mono',Consolas,monospace;color:#d0d0d0}" +
      "</style>" +
      "<h1>" + esc2(title) + "</h1><p class=sub>" + rows.length + " rows · click a header to sort · filter below</p>" +
      '<input id="f" placeholder="Filter rows…" oninput="filter()">' +
      "<table id=t><thead><tr>" + head + "</tr></thead><tbody>" + bodyRows + "</tbody></table>" +
      "<script>" +
      "var t=document.getElementById('t');" +
      "function filter(){var q=document.getElementById('f').value.toLowerCase();" +
      "Array.prototype.forEach.call(t.tBodies[0].rows,function(r){r.style.display=r.textContent.toLowerCase().indexOf(q)>=0?'':'none';});}" +
      "Array.prototype.forEach.call(t.tHead.rows[0].cells,function(th,i){th.onclick=function(){" +
      "var rows=Array.prototype.slice.call(t.tBodies[0].rows);var asc=th._asc=!th._asc;" +
      "rows.sort(function(a,b){var x=a.cells[i].textContent,y=b.cells[i].textContent;" +
      "var nx=parseFloat(x),ny=parseFloat(y);var c=(!isNaN(nx)&&!isNaN(ny))?nx-ny:x.localeCompare(y);return asc?c:-c;});" +
      "rows.forEach(function(r){t.tBodies[0].appendChild(r);});};});" +
      "<\/script>";
  }

  /* ================= tab: Overview ================= */
  function tabOverview(container) {
    container.appendChild(note("Loading list metadata…"));
    spGet(listApi() + "?$select=Title,BaseType,ItemCount,EnableVersioning,MajorVersionLimit,EnableAttachments,HasUniqueRoleAssignments")
      .then(function (d) {
        ctx.listTitle = d.Title || ctx.listTitle;
        container.innerHTML = "";
        container.appendChild(kv([
          ["Title", d.Title],
          ["Type", d.BaseType === 1 ? "Document Library" : "List"],
          ["Items", d.ItemCount],
          ["Versioning", d.EnableVersioning ? "On (limit " + (d.MajorVersionLimit || "∞") + ")" : "Off"],
          ["Attachments", d.EnableAttachments ? "Enabled" : "Disabled"],
          ["Permissions", d.HasUniqueRoleAssignments ? "Unique" : "Inherited"]
        ]));
      })
      .catch(function (e) { container.innerHTML = ""; container.appendChild(note("Could not read list. " + e.message)); });
  }

  /* ================= tab: Columns ================= */
  function tabColumns(container) {
    container.appendChild(note("Loading columns…"));
    spGetAll(listApi() + "/fields?$select=Title,InternalName,TypeAsString,Hidden,Required,ReadOnlyField&$top=500")
      .then(function (fields) {
        fields = fields.filter(function (f) { return !f.Hidden; });
        container.innerHTML = "";
        var bar = el("div", "row");
        var search = el("input"); search.className = "search"; search.placeholder = "Filter columns…";
        bar.appendChild(search);
        bar.appendChild(btn("JSON", function () {
          downloadBlob(JSON.stringify(fields, null, 2), "columns.json", "application/json");
        }));
        bar.appendChild(btn("CSV", function () {
          downloadBlob(toCsv(fields, ["Title", "InternalName", "TypeAsString", "Required", "ReadOnlyField"]), "columns.csv", "text/csv");
        }));
        container.appendChild(bar);
        var mount = el("div");
        container.appendChild(mount);
        function draw(q) {
          mount.innerHTML = "";
          fields.filter(function (f) {
            return !q || (f.Title + " " + f.InternalName + " " + f.TypeAsString).toLowerCase().indexOf(q) >= 0;
          }).forEach(function (f) {
            var r = el("div", "field");
            r.innerHTML = '<b>' + esc(f.Title) + '</b>' +
              '<code title="Click to copy">' + esc(f.InternalName) + '</code>' +
              '<span class="tag">' + esc(f.TypeAsString) + '</span>' +
              (f.Required ? '<span class="tag req">required</span>' : '');
            r.querySelector("code").onclick = function () { copyText(f.InternalName); };
            mount.appendChild(r);
          });
        }
        search.oninput = function () { draw(search.value.toLowerCase()); };
        draw("");
      })
      .catch(function (e) { container.innerHTML = ""; container.appendChild(note("Error: " + e.message)); });
  }

  /* ================= tab: Views ================= */
  function tabViews(container) {
    container.appendChild(note("Loading views…"));
    spGetAll(listApi() + "/views?$select=Title,RowLimit,ViewQuery,DefaultView&$top=200")
      .then(function (views) {
        container.innerHTML = "";
        views.forEach(function (v) {
          var r = el("div", "field col");
          r.innerHTML = '<b>' + esc(v.Title) + (v.DefaultView ? ' <span class="tag">default</span>' : '') + '</b>' +
            '<span class="muted">Row limit: ' + esc(v.RowLimit) + '</span>' +
            '<pre>' + esc(v.ViewQuery || "(no CAML query)") + '</pre>';
          container.appendChild(r);
        });
        if (!views.length) container.appendChild(note("No views."));
      })
      .catch(function (e) { container.innerHTML = ""; container.appendChild(note("Error: " + e.message)); });
  }

  /* ================= tab: Recycle Bin ================= */
  function tabRecycle(container) {
    container.appendChild(note("Loading first-stage recycle bin…"));
    var folderHint = (ctx.webServerRel + "/Lists/" + (ctx.listTitle || "")).toLowerCase();
    spGetAll("/_api/web/recyclebin?$select=Id,Title,DirName,ItemType,DeletedDate&$top=200")
      .then(function (all) {
        var items = all.filter(function (r) {
          return (r.DirName || "").toLowerCase().indexOf((ctx.listTitle || "___nolist___").toLowerCase()) >= 0
            || (r.DirName || "").toLowerCase().indexOf(folderHint) >= 0;
        });
        container.innerHTML = "";
        if (!items.length) { container.appendChild(note("No matching deleted items for this list.")); return; }
        var checks = [];
        items.forEach(function (r) {
          var row = el("div", "field");
          var cb = el("input"); cb.type = "checkbox"; cb.dataset.id = r.Id;
          checks.push(cb);
          var lbl = el("span"); lbl.innerHTML = '<b>' + esc(r.Title) + '</b> <span class="muted">' + esc(r.DeletedDate) + '</span>';
          row.appendChild(cb); row.appendChild(lbl);
          container.appendChild(row);
        });
        container.appendChild(btn("Restore checked", function () {
          var ids = checks.filter(function (c) { return c.checked; }).map(function (c) { return c.dataset.id; });
          if (!ids.length) { log("Nothing checked."); return; }
          if (!confirm("Restore " + ids.length + " item(s) from the recycle bin?")) return;
          restoreMany(ids, 0);
        }));
      })
      .catch(function (e) { container.innerHTML = ""; container.appendChild(note("Error: " + e.message)); });
  }
  function restoreMany(ids, i) {
    if (i >= ids.length) { log("Restore complete."); selectTab("recycle"); return; }
    spPost("/_api/web/recyclebin('" + ids[i] + "')/restore()")
      .then(function () { log("Restored " + (i + 1) + "/" + ids.length); restoreMany(ids, i + 1); })
      .catch(function (e) { log("Restore failed: " + e.message); });
  }

  /* ================= tab: Export ================= */
  function tabExport(container) {
    container.appendChild(note("Snapshot every item in this list as rollback insurance before a bulk edit."));
    var bar = el("div", "row");
    bar.appendChild(btn("Export JSON", function () { doExport("json"); }));
    bar.appendChild(btn("Export CSV", function () { doExport("csv"); }));
    container.appendChild(bar);
    var out = el("div"); container.appendChild(out);
    function doExport(fmt) {
      out.innerHTML = ""; out.appendChild(note("Reading all items…"));
      spGetAll(listApi() + "/items?$top=2000")
        .then(function (items) {
          out.innerHTML = "";
          if (fmt === "json") {
            downloadBlob(JSON.stringify(items, null, 2), "list-export.json", "application/json");
          } else {
            var cols = items.length ? Object.keys(items[0]).filter(function (k) { return k.indexOf("odata") < 0; }) : [];
            downloadBlob(toCsv(items, cols), "list-export.csv", "text/csv");
          }
          out.appendChild(note("Exported " + items.length + " items."));
        })
        .catch(function (e) { out.innerHTML = ""; out.appendChild(note("Error: " + e.message)); });
    }
  }

  /* ================= tab: Versions ================= */
  var REAPPLY_TYPES = { Text: 1, Note: 1, Number: 1, Currency: 1, DateTime: 1, Choice: 1, Boolean: 1, Integer: 1 };
  function tabVersions(container) {
    var bar = el("div", "row");
    var idIn = el("input"); idIn.className = "search"; idIn.placeholder = "Item ID"; idIn.type = "number";
    bar.appendChild(idIn);
    bar.appendChild(btn("Load history", function () { loadVersions(idIn.value); }));
    container.appendChild(bar);
    var mount = el("div"); container.appendChild(mount);

    function loadVersions(id) {
      if (!id) { log("Enter an item ID."); return; }
      mount.innerHTML = ""; mount.appendChild(note("Loading versions…"));
      spGetAll(listApi() + "/items(" + id + ")/versions?$top=200")
        .then(function (versions) {
          mount.innerHTML = "";
          if (!versions.length) { mount.appendChild(note("No versions.")); return; }
          versions.forEach(function (v) {
            var r = el("div", "field col");
            r.innerHTML = '<b>v' + esc(v.VersionLabel) + '</b>' +
              '<span class="muted">' + esc(v.Modified || "") + " · " + esc((v.Editor && v.Editor.LookupValue) || "") + '</span>';
            r.appendChild(btn("Re-apply editable fields", function () {
              if (!confirm("Re-apply v" + v.VersionLabel + " to item " + id + "?\nOnly simple text/number/date/choice/yes-no fields are written; lookups, people, and metadata are skipped.")) return;
              reapply(id, v);
            }));
            mount.appendChild(r);
          });
        })
        .catch(function (e) { mount.innerHTML = ""; mount.appendChild(note("Error: " + e.message)); });
    }
    function reapply(id, version) {
      // Fetch field types once, then write back only the safe subset.
      spGetAll(listApi() + "/fields?$select=InternalName,TypeAsString,ReadOnlyField,Hidden&$top=500")
        .then(function (fields) {
          var body = {};
          fields.forEach(function (f) {
            if (f.ReadOnlyField || f.Hidden) return;
            if (!REAPPLY_TYPES[f.TypeAsString]) return;
            if (Object.prototype.hasOwnProperty.call(version, f.InternalName)) {
              body[f.InternalName] = version[f.InternalName];
            }
          });
          var count = Object.keys(body).length;
          if (!count) { log("No re-appliable fields found."); return; }
          spPost(listApi() + "/items(" + id + ")", body, { merge: true })
            .then(function () { log("Re-applied " + count + " field(s) to item " + id); })
            .catch(function (e) { log("Re-apply failed: " + e.message); });
        });
    }
  }

  /* ================= shared UI utilities ================= */
  function el(tag, cls) { var n = document.createElement(tag); if (cls) n.className = cls; return n; }
  function btn(label, fn) { var b = el("button", "btn"); b.textContent = label; b.onclick = fn; return b; }
  function note(text) { var n = el("div", "note"); n.textContent = text; return n; }
  function kv(pairs) {
    var box = el("div", "kv");
    pairs.forEach(function (p) {
      var r = el("div", "kv-row");
      r.innerHTML = '<span class="k">' + esc(p[0]) + '</span><span class="v">' + esc(p[1] == null ? "—" : p[1]) + '</span>';
      box.appendChild(r);
    });
    return box;
  }
  var logTimer;
  function log(msg) {
    logEl.textContent = msg; logEl.classList.add("show");
    clearTimeout(logTimer); logTimer = setTimeout(function () { logEl.classList.remove("show"); }, 2600);
  }
  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { log("Copied"); }, fallback);
    } else fallback();
    function fallback() {
      var t = document.createElement("textarea"); t.value = text;
      t.style.position = "fixed"; t.style.opacity = "0";
      root.appendChild(t); t.select();
      try { document.execCommand("copy"); log("Copied"); } catch (e) { log("Copy failed"); }
      root.removeChild(t);
    }
  }
  function toCsv(rows, cols) {
    var lines = [cols.join(",")];
    rows.forEach(function (r) {
      lines.push(cols.map(function (c) {
        var v = r[c]; if (v == null) v = "";
        v = String(typeof v === "object" ? JSON.stringify(v) : v);
        return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
      }).join(","));
    });
    return lines.join("\r\n");
  }
  function downloadBlob(text, name, type) {
    var blob = new Blob([text], { type: type });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    root.appendChild(a); a.click(); root.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  function pickFile(cb) {
    var inp = document.createElement("input"); inp.type = "file"; inp.accept = "application/json";
    inp.onchange = function () {
      var f = inp.files[0]; if (!f) return;
      var r = new FileReader(); r.onload = function () { cb(r.result); }; r.readAsText(f);
    };
    inp.click();
  }

  /* ================= public toggle + boot ================= */
  var api = {
    visible: true,
    toggle: function () {
      api.visible = !api.visible;
      host.style.display = api.visible ? "block" : "none";
    }
  };
  window.__SPCP = api;

  loadLinks();
  selectTab("links");
  if (!storageOk) log("localStorage blocked — links won't persist here.");

  /* ================= styles ================= */
  function cssText() {
    // TrueCrimeDev design system, scoped to the shadow root. Depth via borders,
    // not shadows; accents (blue/green/amber/red/purple/cyan) carry meaning only.
    return "" +
    ":host{all:initial;" +
      "--tc-brand:#5B9FEF;--tc-brand-fg:#0f0f0f;" +
      "--tc-bg-chrome:#141414;--tc-bg-card:#121212;--tc-bg-elevated:#1a1a1a;--tc-bg-elevated-high:#202020;" +
      "--tc-bg-hover:#252525;--tc-bg-active:#2a2a2a;" +
      "--tc-border-subtle:#232323;--tc-border:#303030;--tc-border-hover:#505050;--tc-border-strong:#707070;" +
      "--tc-fg:#ffffff;--tc-fg-secondary:#d0d0d0;--tc-fg-secondary-dim:#b0b0b0;--tc-fg-tertiary:#a0a0a0;" +
      "--tc-fg-muted:#808080;--tc-fg-muted-dim:#707070;--tc-fg-placeholder:#606060;" +
      "--tc-blue:#5B9FEF;--tc-green:#7BC96F;--tc-amber:#F59E42;--tc-red:#DC3545;--tc-purple:#A855F7;--tc-cyan:#22D3EE;" +
      "--tc-ring-brand:0 0 0 3px #5B9FEF30;" +
      "--tc-font-sans:'DM Sans',system-ui,'Segoe UI',Roboto,Arial,sans-serif;" +
      "--tc-font-mono:'JetBrains Mono','Cascadia Code',Consolas,ui-monospace,monospace;" +
      "--tc-ease:cubic-bezier(0.2,0,0,1);--tc-dur:180ms;--tc-dur-fast:120ms}" +
    "*{box-sizing:border-box;font-family:var(--tc-font-sans)}" +
    ".panel{width:420px;max-height:78vh;display:flex;flex-direction:column;background:var(--tc-bg-card);color:var(--tc-fg);" +
      "border:1px solid var(--tc-border);border-radius:12px;box-shadow:0 12px 32px rgba(0,0,0,.5);font-size:13px;overflow:hidden}" +
    ".hdr{display:flex;align-items:center;gap:8px;padding:10px 12px;background:var(--tc-bg-chrome);" +
      "border-bottom:1px solid var(--tc-border);cursor:move;user-select:none}" +
    ".hdr .title{font-family:var(--tc-font-mono);font-weight:500;letter-spacing:-0.01em}" +
    ".hdr .title b{color:var(--tc-brand);font-weight:500}" +
    ".hdr .ctx{font-size:11px;color:var(--tc-fg-muted);border:1px solid var(--tc-border);border-radius:999px;padding:1px 8px}" +
    ".hdr .spacer{flex:1}" +
    ".hdr .x{background:transparent;border:0;color:var(--tc-fg-muted);font-size:18px;cursor:pointer;line-height:1;transition:color var(--tc-dur-fast) var(--tc-ease)}" +
    ".hdr .x:hover{color:var(--tc-fg)}" +
    ".tabs{display:flex;flex-wrap:wrap;gap:4px;padding:8px;border-bottom:1px solid var(--tc-border);background:var(--tc-bg-chrome)}" +
    ".tab{background:var(--tc-bg-elevated);border:1px solid var(--tc-border);color:var(--tc-fg-tertiary);border-radius:8px;" +
      "padding:5px 10px;font-size:12px;font-weight:600;cursor:pointer;" +
      "transition:background var(--tc-dur) var(--tc-ease),border-color var(--tc-dur) var(--tc-ease),color var(--tc-dur) var(--tc-ease)}" +
    ".tab:hover{background:var(--tc-bg-hover);border-color:var(--tc-border-hover);color:var(--tc-fg)}" +
    ".tab.active{background:var(--tc-brand);border-color:var(--tc-brand);color:var(--tc-brand-fg)}" +
    ".tab.disabled{opacity:.4;cursor:not-allowed}" +
    ".body{padding:12px;overflow:auto;flex:1;display:flex;flex-direction:column;gap:12px}" +
    ".row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}" +
    ".btn{background:var(--tc-bg-elevated);border:1px solid var(--tc-border);color:var(--tc-fg-secondary);border-radius:8px;" +
      "padding:6px 11px;font-size:12px;font-weight:600;cursor:pointer;" +
      "transition:background var(--tc-dur) var(--tc-ease),border-color var(--tc-dur) var(--tc-ease),color var(--tc-dur) var(--tc-ease)}" +
    ".btn:hover{background:var(--tc-bg-hover);border-color:var(--tc-border-hover);color:var(--tc-fg)}" +
    ".btn:active{background:var(--tc-bg-active)}" +
    "textarea{width:100%;min-height:64px;resize:vertical;background:var(--tc-bg-elevated-high);color:var(--tc-fg);border:1px solid var(--tc-border);" +
      "border-radius:8px;padding:8px;font-family:var(--tc-font-mono);font-size:12px;line-height:1.5}" +
    "input.search,input[type=number]{width:100%;background:var(--tc-bg-elevated-high);color:var(--tc-fg);border:1px solid var(--tc-border);border-radius:8px;padding:7px 9px;font-size:12px}" +
    "textarea::placeholder,input::placeholder{color:var(--tc-fg-placeholder)}" +
    "textarea:focus,input:focus{outline:none;border-color:var(--tc-brand);box-shadow:var(--tc-ring-brand)}" +
    ".note{color:var(--tc-fg-muted);font-size:12px;border:1px dashed var(--tc-border);border-radius:8px;padding:10px}" +
    ".grp-h{font-size:11px;text-transform:uppercase;letter-spacing:1.2px;font-weight:600;color:var(--tc-fg-muted-dim);margin:8px 0 4px}" +
    ".lcard{--a:var(--tc-cyan);background:var(--tc-bg-card);border:1px solid var(--tc-border);border-radius:12px;padding:10px 12px;margin-bottom:8px;" +
      "transition:background var(--tc-dur) var(--tc-ease),border-color var(--tc-dur) var(--tc-ease)}" +
    ".lcard:hover{background:var(--tc-bg-elevated);border-color:var(--tc-border-hover)}" +
    ".lcard[data-t=LIST]{--a:var(--tc-blue)}.lcard[data-t=LIBRARY]{--a:var(--tc-green)}" +
    ".lcard[data-t=VIEW]{--a:var(--tc-amber)}.lcard[data-t=PAGE]{--a:var(--tc-purple)}" +
    ".lcard[data-t=SYSTEM]{--a:var(--tc-red)}.lcard[data-t=LINK]{--a:var(--tc-cyan)}" +
    ".lname{font-weight:600;font-size:14px;color:var(--tc-fg);outline:none;word-break:break-word;letter-spacing:-0.01em}" +
    ".lname[contenteditable=true]{background:var(--tc-bg-elevated-high);border-radius:6px;padding:1px 4px}" +
    ".lmeta{margin:6px 0}.lacts{display:flex;gap:6px}" +
    ".lacts .btn{padding:4px 9px}" +
    ".tag{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;letter-spacing:.4px;padding:2px 9px;" +
      "border-radius:999px;border:1px solid var(--a,var(--tc-fg-muted));color:var(--tc-fg-secondary-dim);line-height:1.6}" +
    ".tag::before{content:'';width:6px;height:6px;border-radius:999px;background:var(--a,var(--tc-fg-muted));flex:none}" +
    ".tag.req{--a:var(--tc-red);color:var(--tc-red)}" +
    ".kv{border:1px solid var(--tc-border);border-radius:12px;overflow:hidden}" +
    ".kv-row{display:flex;justify-content:space-between;gap:12px;padding:8px 12px;border-bottom:1px solid var(--tc-border-subtle)}" +
    ".kv-row:last-child{border-bottom:0}.kv .k{color:var(--tc-fg-muted)}.kv .v{font-weight:500;font-family:var(--tc-font-mono);color:var(--tc-fg)}" +
    ".field{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:var(--tc-bg-card);border:1px solid var(--tc-border);" +
      "border-radius:8px;padding:8px 11px;margin-bottom:6px;transition:border-color var(--tc-dur) var(--tc-ease)}" +
    ".field:hover{border-color:var(--tc-border-hover)}" +
    ".field.col{flex-direction:column;align-items:stretch}" +
    ".field b{font-weight:600;color:var(--tc-fg)}" +
    ".field code{background:var(--tc-bg-elevated-high);border:1px solid var(--tc-border);border-radius:6px;padding:1px 6px;cursor:pointer;" +
      "font-family:var(--tc-font-mono);font-size:11px;color:var(--tc-fg-secondary-dim);transition:border-color var(--tc-dur) var(--tc-ease)}" +
    ".field code:hover{border-color:var(--tc-brand);color:var(--tc-fg)}" +
    ".field .muted,.muted{color:var(--tc-fg-muted);font-size:11px}" +
    ".field pre{background:var(--tc-bg-elevated-high);border:1px solid var(--tc-border);border-radius:6px;padding:8px;overflow:auto;" +
      "font-family:var(--tc-font-mono);font-size:11px;color:var(--tc-fg-secondary-dim);margin:4px 0 0}" +
    ".btn.is-disabled{opacity:.4;cursor:not-allowed;pointer-events:none}" +
    ".tbl-scroll{overflow:auto;max-height:46vh;border:1px solid var(--tc-border);border-radius:8px}" +
    ".tbl{border-collapse:collapse;width:100%;font-size:12px}" +
    ".tbl th{position:sticky;top:0;background:var(--tc-bg-chrome);color:var(--tc-fg-tertiary);text-align:left;" +
      "font-size:10px;text-transform:uppercase;letter-spacing:1px;padding:6px 9px;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:1px solid var(--tc-border)}" +
    ".tbl th:hover{color:var(--tc-fg)}" +
    ".tbl td{padding:6px 9px;border-bottom:1px solid var(--tc-border-subtle);color:var(--tc-fg-secondary-dim);" +
      "font-family:var(--tc-font-mono);max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}" +
    ".tbl tr.clickable{cursor:pointer}.tbl tr.clickable:hover td{background:var(--tc-bg-hover)}" +
    ".tbl a.tlink{color:var(--tc-brand);text-decoration:none}.tbl a.tlink:hover{text-decoration:underline}" +
    ".log{padding:0 12px;max-height:0;overflow:hidden;color:var(--tc-green);font-size:12px;transition:max-height var(--tc-dur) var(--tc-ease),padding var(--tc-dur) var(--tc-ease)}" +
    ".log.show{max-height:40px;padding:8px 12px;border-top:1px solid var(--tc-border)}";
  }
})();
