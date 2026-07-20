/*
 * read-list.test.js - automated tests for browser/read-list.js
 * -------------------------------------------------------------
 * Plain Node, zero dependencies, no jsdom. The browser reader is an IIFE that
 * expects browser globals, so we hand-roll them (window, fetch, console,
 * document, Blob, URL, location, setTimeout, JSON) and evaluate the source with
 * `new Function`, letting its `window.SPLA = ...` assignment land on our fake
 * window. Every scenario gets a FRESH load, because the reader keeps the picked
 * export folder in a module-level `outDir` closure variable.
 *
 * `JSON` is injected too (it shadows the real global inside the evaluated
 * source) so a serialization failure can be simulated without corrupting the
 * fixtures.
 *
 * The fake directory handle is a REAL in-memory file model - a name -> entry
 * map where getFileHandle({create:true}) creates the entry immediately (as the
 * browser does), a writable stream only commits its buffer on close(), abort()
 * rolls the entry back, and removeEntry() deletes it. That is what lets the
 * folder scenarios observe leftover 0-byte files and orphaned streams instead
 * of merely observing that some function was called.
 *
 * Run it:
 *
 *   Windows PowerShell:   node .\test\read-list.test.js
 *   (from C:\SP\sharepoint-list-agent)
 *
 * Exits 0 when every check passes, 1 otherwise.
 */
"use strict";

var fs = require("fs");
var path = require("path");

var SRC_PATH = process.env.SPLA_SRC || path.join(__dirname, "..", "browser", "read-list.js");
var SRC = fs.readFileSync(SRC_PATH, "utf8");

var WEB = "https://contoso.sharepoint.com/sites/whd";
var LIST = "My List";
var LIST_ID = "11111111-2222-3333-4444-555555555555";

// A different title that sanitizes to the SAME stem as LIST ("My_List").
var LIST_TWIN = "My@List";
var LIST_TWIN_ID = "77777777-8888-9999-aaaa-bbbbbbbbbbbb";

var LIST_IDS = {};
LIST_IDS[LIST] = LIST_ID;
LIST_IDS[LIST_TWIN] = LIST_TWIN_ID;

/* ------------------------------------------------------------------ *
 * Fixture: the field set returned by /fields
 * ------------------------------------------------------------------ */

// Kept: CanBeDeleted, or an internal name in the reader's SYS_KEEP set.
var KEPT_FIELDS = [
  { Title: "Title",       InternalName: "Title",       TypeAsString: "Text",                Hidden: false, CanBeDeleted: false },
  { Title: "Created",     InternalName: "Created",     TypeAsString: "DateTime",            Hidden: false, CanBeDeleted: false },
  { Title: "Modified",    InternalName: "Modified",    TypeAsString: "DateTime",            Hidden: false, CanBeDeleted: false },
  { Title: "Created By",  InternalName: "Author",      TypeAsString: "User",                Hidden: false, CanBeDeleted: false },
  { Title: "Name",        InternalName: "FileLeafRef", TypeAsString: "File",                Hidden: false, CanBeDeleted: false },
  { Title: "Active",      InternalName: "Active",      TypeAsString: "Boolean",             Hidden: false, CanBeDeleted: true },
  { Title: "Doc Link",    InternalName: "DocLink",     TypeAsString: "URL",                 Hidden: false, CanBeDeleted: true },
  { Title: "Categories",  InternalName: "Categories",  TypeAsString: "MultiChoice",         Hidden: false, CanBeDeleted: true },
  { Title: "Region",      InternalName: "Region",      TypeAsString: "TaxonomyFieldType",   Hidden: false, CanBeDeleted: true },
  { Title: "Where",       InternalName: "Where",       TypeAsString: "Geolocation",         Hidden: false, CanBeDeleted: true },
  { Title: "Notes",       InternalName: "Notes",       TypeAsString: "Note",                Hidden: false, CanBeDeleted: true },
  { Title: "Assigned To", InternalName: "AssignedTo",  TypeAsString: "User",                Hidden: false, CanBeDeleted: true },
  { Title: "Case",        InternalName: "CaseLink",    TypeAsString: "Lookup",              Hidden: false, CanBeDeleted: true, LookupField: "CaseNumber" }
];

// Every one of these must be filtered out, for a different documented reason.
var DROPPED_FIELDS = [
  { Title: "Secret",              InternalName: "SecretField",      TypeAsString: "Text",                   Hidden: true,  CanBeDeleted: true },
  { Title: "UI Version",          InternalName: "_UIVersionString", TypeAsString: "Text",                   Hidden: false, CanBeDeleted: true },
  { Title: "Doc Icon",            InternalName: "DocIcon",          TypeAsString: "Computed",               Hidden: false, CanBeDeleted: true },
  { Title: "Enterprise Keywords", InternalName: "TaxKeyword",       TypeAsString: "TaxonomyFieldTypeMulti", Hidden: false, CanBeDeleted: true },
  { Title: "Attachments",         InternalName: "Attachments",      TypeAsString: "Attachments",            Hidden: false, CanBeDeleted: true },
  { Title: "Content Type",        InternalName: "ContentType",      TypeAsString: "ContentType",            Hidden: false, CanBeDeleted: true },
  { Title: "GUID",                InternalName: "GUID",             TypeAsString: "Guid",                   Hidden: false, CanBeDeleted: true },
  // Not deletable and not in SYS_KEEP -> dropped.
  { Title: "Order",               InternalName: "Order",            TypeAsString: "Number",                 Hidden: false, CanBeDeleted: false }
];

var ALL_FIELDS = KEPT_FIELDS.concat(DROPPED_FIELDS);

var KEPT_INTERNALS = KEPT_FIELDS.map(function (f) { return f.InternalName; });
var KEPT_TITLES = KEPT_FIELDS.map(function (f) { return f.Title; });

// Items are built scalars-first, then users, then lookups - so their key order
// is NOT the raw field order that `columns` preserves.
function titlesOfType(types) {
  return KEPT_FIELDS.filter(function (f) { return types.indexOf(f.TypeAsString) !== -1; })
    .map(function (f) { return f.Title; });
}
var USER_TITLES = titlesOfType(["User", "UserMulti"]);
var LOOKUP_TITLES = titlesOfType(["Lookup", "LookupMulti"]);
var SCALAR_TITLES = KEPT_TITLES.filter(function (t) {
  return USER_TITLES.indexOf(t) === -1 && LOOKUP_TITLES.indexOf(t) === -1;
});
var ITEM_KEY_ORDER = ["Id"].concat(SCALAR_TITLES, USER_TITLES, LOOKUP_TITLES);

/* The six keys the export has always carried, plus the `expanded` provenance
   flag. Order of `expanded` within the object is deliberately NOT pinned; the
   original six must keep their relative order. */
var BASE_KEYS = ["web", "list", "listId", "exportedAt", "columns", "items"];
var SHAPE_KEYS = BASE_KEYS.concat(["expanded"]).slice().sort();

/* ------------------------------------------------------------------ *
 * Fixture: /items rows (expanded query, two pages via odata.nextLink)
 * ------------------------------------------------------------------ */

var ITEMS_PAGE1 = [
  {
    Id: 1,
    Title: "Alpha",
    Created: "2026-01-02T00:00:00Z",
    Modified: "2026-01-03T00:00:00Z",
    FileLeafRef: "alpha.docx",
    Active: true,
    DocLink: { Description: "Case file", Url: "https://contoso.sharepoint.com/f/a.pdf" },
    Categories: ["Wage", "Hour"],
    Region: { Label: "West", TermGuid: "aaaa-bbbb" },
    Where: { Latitude: 38.9, Longitude: -77.03 },
    Notes: { unexpected: "shape", n: 1 },
    Author: { Title: "Ada Lovelace", EMail: "ada@example.gov" },
    AssignedTo: { Title: "Jane Doe", EMail: "jane@example.gov" },
    CaseLink: { Id: 3, CaseNumber: "WHD-2024-001", Title: "TITLE-MUST-NOT-WIN" }
  },
  {
    Id: 2,
    Title: "Beta",
    Created: "2026-02-02T00:00:00Z",
    Modified: "2026-02-03T00:00:00Z",
    FileLeafRef: "beta.docx",
    Active: false,
    DocLink: { Description: "", Url: "https://contoso.sharepoint.com/f/b.pdf" },
    Categories: { results: ["Solo"] },
    Region: null,
    Where: null,
    Notes: "plain text",
    Author: { Title: "Bob Barker", EMail: "bob@example.gov" },
    AssignedTo: null,
    // Show-field present but null, and no Title either -> must degrade to "".
    CaseLink: { Id: 9, CaseNumber: null }
  }
];

var ITEMS_PAGE2 = [
  {
    Id: 3,
    Title: "Gamma",
    Created: "2026-03-02T00:00:00Z",
    Modified: "2026-03-03T00:00:00Z",
    FileLeafRef: "gamma.docx",
    Active: true,
    DocLink: null,
    Categories: { results: [] },
    Region: { Label: "East", TermGuid: "cccc-dddd" },
    Where: { Latitude: 0, Longitude: 0 },
    Notes: "",
    Author: { Title: "Cy Coder", EMail: "cy@example.gov" },
    AssignedTo: { results: [{ Title: "X One", EMail: "x@example.gov" }, { Title: "Y Two", EMail: "y@example.gov" }] },
    CaseLink: { results: [{ CaseNumber: "WHD-A" }, { CaseNumber: "WHD-B" }] }
  }
];

/* Rows returned by the FLAT fallback query (no $expand). */
var ITEMS_FLAT = [
  {
    Id: 1,
    Title: "Alpha",
    Created: "2026-01-02T00:00:00Z",
    Modified: "2026-01-03T00:00:00Z",
    FileLeafRef: "alpha.docx",
    Active: true,
    DocLink: { Description: "Case file", Url: "https://contoso.sharepoint.com/f/a.pdf" },
    Categories: ["Wage", "Hour"],
    Region: { Label: "West" },
    Where: { Latitude: 38.9, Longitude: -77.03 },
    Notes: "n/a",
    AuthorId: 7,
    AssignedToId: 12,
    CaseLinkId: 3
  },
  {
    Id: 2,
    Title: "Beta",
    Created: "2026-02-02T00:00:00Z",
    Modified: "2026-02-03T00:00:00Z",
    FileLeafRef: "beta.docx",
    Active: false,
    DocLink: null,
    Categories: [],
    Region: null,
    Where: null,
    Notes: "n/a",
    AuthorId: 7,
    AssignedToId: null,
    CaseLinkId: null
  }
];

var EXPANDED_ERR_BODY = "The expression \"AssignedTo/EMail\" is not recognized.";
var FLAT_ERR_BODY = "Field or property 'CaseLinkId' does not exist.";

/* ------------------------------------------------------------------ *
 * Fake browser environment
 * ------------------------------------------------------------------ */

function jsonRes(body) {
  return {
    status: 200,
    ok: true,
    headers: { get: function () { return null; } },
    json: function () { return Promise.resolve(body); },
    text: function () { return Promise.resolve(JSON.stringify(body)); }
  };
}

function errRes(status, body, headers) {
  var h = headers || {};
  return {
    status: status,
    ok: false,
    headers: {
      get: function (k) {
        var v = h[String(k).toLowerCase()];
        return v == null ? null : String(v);
      }
    },
    json: function () { return Promise.resolve({}); },
    text: function () { return Promise.resolve(body); }
  };
}

/** Recover the list title from a getByTitle('...') URL, undoing OData escaping. */
function titleFromUrl(url) {
  var m = /getByTitle\('([^']*(?:''[^']*)*)'\)/.exec(url);
  if (!m) return null;
  var raw;
  try { raw = decodeURIComponent(m[1]); } catch (e) { raw = m[1]; }
  return raw.replace(/''/g, "'");
}

/**
 * Build a complete fake browser global set.
 *
 * opts:
 *   throttleFirstItems  first /items request answers 429 + Retry-After
 *   retryAfter          value of that header (seconds, default "1")
 *   expandedFails       any /items request carrying $expand= answers 400
 *   flatFails           any /items request without $expand= answers 400
 *   noPicker            omit window.showDirectoryPicker entirely
 *   pickerRejects       showDirectoryPicker rejects (user cancelled)
 *   folderWriteThrows   every write() into the folder throws
 *   folderWriteFailTimes  only the first N write() calls throw (transient)
 *   stringifyThrows     JSON.stringify of the export payload throws
 *   noContext           omit window._spPageContextInfo
 *   permission          {query, request, queryThrows} -> the directory handle
 *                       exposes queryPermission/requestPermission returning
 *                       those states. OMITTED entirely (the default) models an
 *                       engine that has no permission API at all, which every
 *                       pre-existing scenario relies on.
 *
 * Mutable switches the scenarios flip mid-run: env.failAllWrites,
 * env.failNextWrites, env.stringifyThrows, env.failPicker.
 */
function makeEnv(opts) {
  opts = opts || {};

  var env = {
    opts: opts,
    requests: [],       // every URL fetch() saw, in order
    timers: [],         // every delay (ms) passed to setTimeout
    logs: [],
    warns: [],
    itemsRequests: 0,
    fieldsRequests: 0,
    blobs: [],
    anchors: [],
    appended: [],
    downloads: [],      // {name, href} per anchor click
    objectUrls: 0,
    revoked: 0,
    fsCalls: [],        // ["getFileHandle", name, create] / ["createWritable"] / ["write"] / ["close"] / ["abort"] / ["removeEntry", name]
    writes: [],         // {name, text} per COMMITTED (closed) stream
    writeAttempts: 0,
    files: {},          // the in-memory folder: name -> {content, committed}
    streams: [],        // every writable stream ever handed out
    pickerCalls: 0,
    permCalls: [],      // ["queryPermission"|"requestPermission", mode]
    perm: opts.permission || null,
    // mutable failure switches (tests flip these mid-scenario)
    failAllWrites: !!opts.folderWriteThrows,
    failNextWrites: opts.folderWriteFailTimes || 0,
    stringifyThrows: !!opts.stringifyThrows,
    failPicker: false   // flipped to model a LATER picker call being cancelled
  };

  env.window = {};
  if (!opts.noContext) {
    env.window._spPageContextInfo = { webAbsoluteUrl: WEB + "/" };
  }

  if (!opts.noPicker) {
    env.window.showDirectoryPicker = function (o) {
      env.pickerCalls++;
      env.fsCalls.push(["showDirectoryPicker", (o && o.mode) || null]);
      if (opts.pickerRejects || env.failPicker) return Promise.reject(new Error("The user aborted a request."));
      return Promise.resolve(makeDirHandle(env));
    };
  }

  env.location = { origin: "https://contoso.sharepoint.com", href: WEB + "/Lists/My%20List/AllItems.aspx" };

  env.console = {
    log: function () { env.logs.push(Array.prototype.join.call(arguments, " ")); },
    warn: function () { env.warns.push(Array.prototype.join.call(arguments, " ")); },
    error: function () { env.warns.push("ERROR " + Array.prototype.join.call(arguments, " ")); },
    table: function (rows) { env.logs.push("TABLE " + JSON.stringify(rows)); }
  };

  // setImmediate (not setTimeout) so no real delay elapses and the event loop
  // is never held open by an orphan timer, while still recording the delay the
  // code asked for - that is what proves Retry-After was honoured.
  env.setTimeout = function (fn, ms) {
    env.timers.push(ms);
    return setImmediate(fn);
  };

  // Shadows the real global inside the evaluated source. Serialization of the
  // export payload can be made to fail on demand; everything else (the
  // flatteners' JSON.stringify of unexpected shapes) is untouched.
  env.JSON = {
    parse: function (t, r) { return JSON.parse(t, r); },
    stringify: function (v, r, s) {
      if (env.stringifyThrows && v && typeof v === "object" && v.columns && v.items && v.exportedAt) {
        throw new TypeError("Converting circular structure to JSON");
      }
      return JSON.stringify(v, r, s);
    }
  };

  env.Blob = function Blob(parts, o) {
    this.parts = parts;
    this.type = o && o.type;
    env.blobs.push(this);
  };

  env.URL = {
    createObjectURL: function () { env.objectUrls++; return "blob:mock/" + env.objectUrls; },
    revokeObjectURL: function () { env.revoked++; }
  };

  env.document = {
    createElement: function (tag) {
      var el = {
        tagName: String(tag).toUpperCase(),
        href: null,
        download: null,
        clicks: 0,
        removed: 0,
        click: function () { this.clicks++; env.downloads.push({ name: this.download, href: this.href }); },
        remove: function () { this.removed++; }
      };
      env.anchors.push(el);
      return el;
    },
    body: { appendChild: function (el) { env.appended.push(el); return el; } }
  };

  env.fetch = function (url) { return route(env, opts, String(url)); };

  return env;
}

/**
 * A directory handle backed by a real in-memory file model.
 *
 * Browser fidelity that matters here:
 *   - getFileHandle(name, {create:true}) creates the entry IMMEDIATELY, as a
 *     0-byte file, before anything is written.
 *   - createWritable() buffers; the file only takes the new bytes on close().
 *   - abort() discards the buffer and restores the previous content, but does
 *     NOT remove the entry - only removeEntry() does that.
 * So a write that fails and is merely aborted leaves a 0-byte file behind,
 * which is exactly the corruption these tests must be able to see.
 */
function makeDirHandle(env) {
  function shouldFail() {
    if (env.failNextWrites > 0) { env.failNextWrites--; return true; }
    return !!env.failAllWrites;
  }

  var dir = {
    getFileHandle: function (name, o) {
      var create = (o && o.create) === true;
      env.fsCalls.push(["getFileHandle", name, create]);
      if (!Object.prototype.hasOwnProperty.call(env.files, name)) {
        if (!create) return Promise.reject(new Error("NotFoundError: " + name + " does not exist"));
        env.files[name] = { content: "", committed: false };
      }
      return Promise.resolve({
        name: name,
        createWritable: function () {
          env.fsCalls.push(["createWritable"]);
          var prev = env.files[name] ? env.files[name].content : "";
          var prevCommitted = env.files[name] ? env.files[name].committed : false;
          var st = { name: name, buf: "", closed: false, aborted: false };
          env.streams.push(st);
          return Promise.resolve({
            write: function (text) {
              env.fsCalls.push(["write"]);
              env.writeAttempts++;
              if (st.closed || st.aborted) return Promise.reject(new Error("InvalidStateError: stream is not writable"));
              if (shouldFail()) return Promise.reject(new Error("NotAllowedError: permission lost"));
              st.buf += (typeof text === "string" ? text : String(text));
              return Promise.resolve();
            },
            close: function () {
              env.fsCalls.push(["close"]);
              if (st.aborted) return Promise.reject(new Error("InvalidStateError: stream was aborted"));
              if (st.closed) return Promise.reject(new Error("InvalidStateError: stream already closed"));
              st.closed = true;
              if (!Object.prototype.hasOwnProperty.call(env.files, name)) {
                env.files[name] = { content: "", committed: false };
              }
              env.files[name].content = st.buf;   // commit ONLY on close
              env.files[name].committed = true;
              env.writes.push({ name: name, text: st.buf });
              return Promise.resolve();
            },
            abort: function () {
              env.fsCalls.push(["abort"]);
              if (st.closed) return Promise.reject(new Error("InvalidStateError: stream already closed"));
              st.aborted = true;
              if (Object.prototype.hasOwnProperty.call(env.files, name)) {
                env.files[name].content = prev;   // rollback, entry survives
                env.files[name].committed = prevCommitted;
              }
              return Promise.resolve();
            }
          });
        }
      });
    },
    removeEntry: function (name) {
      env.fsCalls.push(["removeEntry", name]);
      if (!Object.prototype.hasOwnProperty.call(env.files, name)) {
        return Promise.reject(new Error("NotFoundError: " + name + " does not exist"));
      }
      delete env.files[name];
      return Promise.resolve();
    }
  };

  // The permission API is OPTIONAL on a real handle (older engines lack it), so
  // it only exists when a scenario asks for it. Absence is itself a case under
  // test: the reader must treat an un-checkable grant as usable, not as denied.
  if (env.perm) {
    dir.queryPermission = function (o) {
      env.permCalls.push(["queryPermission", (o && o.mode) || null]);
      if (env.perm.queryThrows) return Promise.reject(new Error("SecurityError: permission query blocked"));
      return Promise.resolve(env.perm.query);
    };
    if (env.perm.request !== undefined) {
      dir.requestPermission = function (o) {
        env.permCalls.push(["requestPermission", (o && o.mode) || null]);
        return Promise.resolve(env.perm.request);
      };
    }
  }

  return dir;
}

function route(env, opts, url) {
  env.requests.push(url);

  if (url.indexOf("/fields") !== -1) {
    env.fieldsRequests++;
    return Promise.resolve(jsonRes({ value: ALL_FIELDS }));
  }

  if (url.indexOf("/items") !== -1) {
    env.itemsRequests++;
    var expanded = url.indexOf("$expand=") !== -1;

    if (opts.throttleFirstItems && env.itemsRequests === 1) {
      return Promise.resolve(errRes(429, "throttled", { "retry-after": opts.retryAfter == null ? "1" : opts.retryAfter }));
    }
    if (expanded && opts.expandedFails) return Promise.resolve(errRes(400, EXPANDED_ERR_BODY));
    if (!expanded && opts.flatFails) return Promise.resolve(errRes(400, FLAT_ERR_BODY));
    if (!expanded) return Promise.resolve(jsonRes({ value: ITEMS_FLAT }));

    if (url.indexOf("$skiptoken=PAGE2") !== -1) return Promise.resolve(jsonRes({ value: ITEMS_PAGE2 }));
    return Promise.resolve(jsonRes({ value: ITEMS_PAGE1, "odata.nextLink": url + "&$skiptoken=PAGE2" }));
  }

  if (url.indexOf("$select=Id,Title,ItemCount") !== -1) {
    var t = titleFromUrl(url);
    var id = (t && LIST_IDS[t]) || ("00000000-0000-0000-0000-" + String(Math.abs(hashOf(t || "")) + 1e11).slice(0, 12));
    return Promise.resolve(jsonRes({ Id: id, Title: t || LIST, ItemCount: 3 }));
  }

  if (url.indexOf("/_api/web/lists?") !== -1) {
    return Promise.resolve(jsonRes({
      value: [
        { Title: "Zeta Library", ItemCount: 4, BaseTemplate: 101, Hidden: false },
        { Title: "My List", ItemCount: 3, BaseTemplate: 100, Hidden: false }
      ]
    }));
  }

  return Promise.resolve(errRes(404, "no mock route for " + url));
}

function hashOf(s) {
  var h = 0;
  for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; }
  return h;
}

/** Evaluate the IIFE against a fake env; returns the SPLA it installed. */
function loadSPLA(env) {
  var factory = new Function(
    "window", "fetch", "console", "document", "Blob", "URL", "location", "setTimeout", "JSON",
    SRC
  );
  factory(env.window, env.fetch, env.console, env.document, env.Blob, env.URL, env.location, env.setTimeout, env.JSON);
  if (!env.window.SPLA) throw new Error("read-list.js did not install window.SPLA");
  return env.window.SPLA;
}

/* ------------------------------------------------------------------ *
 * Tiny assertion + reporting harness
 * ------------------------------------------------------------------ */

var results = [];

function record(name, e) {
  if (e) results.push({ name: name, ok: false, err: e && e.message ? e.message : String(e) });
  else results.push({ name: name, ok: true });
}

/**
 * Synchronous check. A thenable return is REJECTED outright: an async assertion
 * body would otherwise be recorded PASS before it ever settled, and its
 * rejection would surface later as an unhandled rejection. Use checkAsync.
 */
function check(name, fn) {
  try {
    var r = fn();
    if (r && typeof r.then === "function") {
      try { r.then(function () {}, function () {}); } catch (ignored) {}
      throw new Error("assertion returned a promise - use the async form (checkAsync)");
    }
    if (r === false) throw new Error("check returned false");
    record(name);
  } catch (e) {
    record(name, e);
  }
}

/** Awaited variant, for assertions that genuinely need to await something. */
async function checkAsync(name, fn) {
  try {
    var r = await fn();
    if (r === false) throw new Error("check returned false");
    record(name);
  } catch (e) {
    record(name, e);
  }
}

function fail(name, msg) { results.push({ name: name, ok: false, err: msg }); }

function assert(cond, msg) { if (!cond) throw new Error(msg || "assertion failed"); }

function eq(actual, expected, msg) {
  if (actual !== expected) {
    throw new Error((msg ? msg + ": " : "") + "expected " + JSON.stringify(expected) + ", got " + JSON.stringify(actual));
  }
}

function deepEq(actual, expected, msg) {
  var a = JSON.stringify(actual), b = JSON.stringify(expected);
  if (a !== b) throw new Error((msg ? msg + ": " : "") + "expected " + b + ", got " + a);
}

function contains(hay, needle, msg) {
  if (String(hay).indexOf(needle) === -1) {
    throw new Error((msg ? msg + ": " : "") + "expected to contain " + JSON.stringify(needle) + " - got " + JSON.stringify(String(hay).slice(0, 400)));
  }
}

function sortedKeys(o) { return Object.keys(o).slice().sort(); }

/** True when `subset` appears in `keys` in that relative order. */
function relativeOrderOk(keys, subset) {
  var last = -1;
  for (var i = 0; i < subset.length; i++) {
    var at = keys.indexOf(subset[i]);
    if (at === -1 || at < last) return false;
    last = at;
  }
  return true;
}

/** Names currently present in the fake export folder. */
function folderNames(env) { return Object.keys(env.files).slice().sort(); }

/**
 * getFileHandle calls split by intent.
 *
 * A read-only getFileHandle(name) - no {create:true} - is a PROBE: it asks
 * "does this entry already exist?" without materializing anything. Only the
 * create:true call actually opens/creates the export file, so counts and
 * filenames are asserted against `gfhCreates` and the probe is tolerated
 * wherever it appears. `gfhAll` is used where the assertion is "the filesystem
 * was not touched AT ALL".
 */
function gfhCreates(env) {
  return env.fsCalls.filter(function (c) { return c[0] === "getFileHandle" && c[2] === true; });
}
function gfhAll(env) {
  return env.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; });
}
/** The fsCall sequence with existence probes filtered out. */
function fsSeq(env) {
  return env.fsCalls
    .filter(function (c) { return !(c[0] === "getFileHandle" && c[2] !== true); })
    .map(function (c) { return c[0]; });
}
/** Names of the permission methods consulted, in order. */
function permNames(env) { return env.permCalls.map(function (c) { return c[0]; }); }

/** Committed text of a folder entry, or null when it never closed. */
function committed(env, name) {
  var f = env.files[name];
  if (!f) throw new Error("no folder entry named " + JSON.stringify(name) + "; have " + JSON.stringify(folderNames(env)));
  return f.committed ? f.content : null;
}

/** Every writable stream must end up either closed or aborted - never orphaned. */
function assertNoOrphanStreams(env) {
  var orphans = env.streams.filter(function (s) { return !s.closed && !s.aborted; });
  assert(orphans.length === 0,
    orphans.length + " writable stream(s) left neither closed nor aborted: " +
    JSON.stringify(orphans.map(function (s) { return s.name; })));
}

/**
 * Yield a full macrotask turn so callbacks the code handed to setTimeout (which
 * this harness maps to setImmediate) actually run. `await` alone only drains
 * microtasks, so deferred work like URL.revokeObjectURL would never fire.
 */
function flush() { return new Promise(function (r) { setImmediate(r); }); }

/** Register a whole group as failed when its scenario blew up. */
function groupFailed(names, err) {
  names.forEach(function (n) { fail(n, "scenario error: " + err); });
}

/**
 * Oracle for the export filename, expressed as a shape rather than a literal:
 * "sp_" + the title sanitized to [A-Za-z0-9._-] + "_" + YYYYMMDD + an OPTIONAL
 * short disambiguator + ".json".
 *
 * The disambiguator's scheme is deliberately NOT pinned - re-implementing the
 * hash here would just re-assert the source against itself. What must hold is
 * that colliding titles end up with different names, which scenario D2 asserts
 * behaviourally.
 */
function nameRe(title, iso) {
  var stem = String(title).replace(/[^A-Za-z0-9._-]+/g, "_").replace(/^_+|_+$/g, "") || "list";
  var day = iso.slice(0, 10).replace(/-/g, "");
  return new RegExp("^sp_" + stem.replace(/[.*+?^${}()|[\]\\-]/g, "\\$&") + "_" + day + "(_[0-9a-z]{4,12})?\\.json$");
}

function assertName(actual, title, iso, msg) {
  var re = nameRe(title, iso);
  assert(re.test(String(actual)),
    (msg ? msg + ": " : "") + JSON.stringify(actual) + " does not match the filename rule " + re);
}

/** The one name the folder is currently storing (throws when that is ambiguous). */
function soleFolderName(env) {
  var n = folderNames(env);
  if (n.length !== 1) throw new Error("expected exactly 1 folder entry, found " + JSON.stringify(n));
  return n[0];
}

/* ------------------------------------------------------------------ *
 * Scenarios
 * ------------------------------------------------------------------ */

async function main() {

  /* ---------- Scenario A: happy path (paging, throttling, filtering,
     flatteners, people, lookups, shape) ---------- */

  var A_NAMES = [
    "paging: nextLink second page is fetched",
    "paging: both pages concatenated in order",
    "throttling: 429 retried and request eventually succeeds",
    "throttling: Retry-After honoured as seconds -> ms",
    "throttling: warning logged naming the status",
    "fields: hidden field excluded",
    "fields: '_'-prefixed internal name excluded",
    "fields: TYPE_SKIP types excluded",
    "fields: TaxonomyFieldTypeMulti excluded",
    "fields: non-deletable non-SYS_KEEP field excluded",
    "fields: FileLeafRef kept",
    "fields: Author/Created/Modified kept",
    "fields: kept set is exactly the expected internals",
    "flatten: Boolean true -> 'Yes'",
    "flatten: Boolean false -> 'No'",
    "flatten: URL object -> 'Description (Url)'",
    "flatten: URL object without Description -> bare Url",
    "flatten: URL null -> ''",
    "flatten: MultiChoice array -> joined",
    "flatten: MultiChoice {results:[...]} -> joined",
    "flatten: MultiChoice empty -> ''",
    "flatten: TaxonomyFieldType -> Label",
    "flatten: TaxonomyFieldType null -> ''",
    "flatten: Geolocation -> 'lat, lon'",
    "flatten: Geolocation zero coords still emitted",
    "flatten: unexpected object -> JSON string",
    "flatten: no raw object survives into items",
    "people: User expands to 'Title <EMail>'",
    "people: Author expands to 'Title <EMail>'",
    "people: null User -> ''",
    "people: multi-value User joined",
    "lookup: expand-select projects LookupField ('/CaseNumber')",
    "lookup: value comes from LookupField, not Title",
    "lookup: null show-value degrades to '' (never an object)",
    "lookup: multi-value lookup joined on show-field",
    "shape: top-level keys are web/list/listId/exportedAt/expanded/columns/items",
    "shape: listId comes from list metadata",
    "shape: columns entries are {name,internal,type}",
    "shape: columns[0] is the synthetic ID counter",
    "shape: expanded flag is true on the expanded path",
    "shape: NO 'degraded' key on the expanded path"
  ];

  try {
    var envA = makeEnv({ throttleFirstItems: true, retryAfter: "1" });
    var splaA = loadSPLA(envA);
    var dataA = await splaA.read(LIST);

    var itemUrls = envA.requests.filter(function (u) { return u.indexOf("/items") !== -1; });
    var expandedUrl = itemUrls.filter(function (u) { return u.indexOf("$expand=") !== -1 && u.indexOf("$skiptoken") === -1; })[0];
    var colInternals = dataA.columns.map(function (c) { return c.internal; });
    var r1 = dataA.items[0], r2 = dataA.items[1], r3 = dataA.items[2];

    check(A_NAMES[0], function () {
      assert(itemUrls.some(function (u) { return u.indexOf("$skiptoken=PAGE2") !== -1; }), "no follow-up request for page 2");
    });
    check(A_NAMES[1], function () {
      eq(dataA.items.length, 3, "row count");
      deepEq(dataA.items.map(function (i) { return i.Id; }), [1, 2, 3], "row ids/order");
    });
    check(A_NAMES[2], function () {
      // 1 throttled + 1 retry (page 1) + 1 page 2 = 3 /items requests
      eq(envA.itemsRequests, 3, "/items request count");
      eq(dataA.items.length, 3, "rows after retry");
    });
    check(A_NAMES[3], function () {
      assert(envA.timers.indexOf(1000) !== -1, "no 1000ms sleep recorded; timers=" + JSON.stringify(envA.timers));
    });
    check(A_NAMES[4], function () {
      assert(envA.warns.some(function (w) { return /Throttled 429/.test(w); }), "no throttle warning: " + JSON.stringify(envA.warns));
    });

    check(A_NAMES[5], function () { assert(colInternals.indexOf("SecretField") === -1, "hidden field leaked"); });
    check(A_NAMES[6], function () { assert(colInternals.indexOf("_UIVersionString") === -1, "_-prefixed field leaked"); });
    check(A_NAMES[7], function () {
      ["DocIcon", "Attachments", "ContentType", "GUID"].forEach(function (n) {
        assert(colInternals.indexOf(n) === -1, n + " (TYPE_SKIP) leaked");
      });
    });
    check(A_NAMES[8], function () { assert(colInternals.indexOf("TaxKeyword") === -1, "TaxonomyFieldTypeMulti leaked"); });
    check(A_NAMES[9], function () { assert(colInternals.indexOf("Order") === -1, "non-deletable non-SYS_KEEP field leaked"); });
    check(A_NAMES[10], function () {
      assert(colInternals.indexOf("FileLeafRef") !== -1, "FileLeafRef dropped");
      assert("Name" in r1, "FileLeafRef column missing from items");
      eq(r1.Name, "alpha.docx");
    });
    check(A_NAMES[11], function () {
      ["Author", "Created", "Modified"].forEach(function (n) {
        assert(colInternals.indexOf(n) !== -1, n + " dropped");
      });
    });
    check(A_NAMES[12], function () {
      deepEq(colInternals, ["Id"].concat(KEPT_INTERNALS), "column internals");
      // Same set of columns on every row, grouped scalars -> users -> lookups.
      deepEq(Object.keys(r1), ITEM_KEY_ORDER, "item keys");
      deepEq(Object.keys(r3), ITEM_KEY_ORDER, "item keys (page 2)");
      deepEq(Object.keys(r1).slice(1).sort(), KEPT_TITLES.slice().sort(), "item columns vs kept fields");
    });

    check(A_NAMES[13], function () { eq(r1.Active, "Yes"); });
    check(A_NAMES[14], function () { eq(r2.Active, "No"); });
    check(A_NAMES[15], function () { eq(r1["Doc Link"], "Case file (https://contoso.sharepoint.com/f/a.pdf)"); });
    check(A_NAMES[16], function () { eq(r2["Doc Link"], "https://contoso.sharepoint.com/f/b.pdf"); });
    check(A_NAMES[17], function () { eq(r3["Doc Link"], ""); });
    check(A_NAMES[18], function () { eq(r1.Categories, "Wage; Hour"); });
    check(A_NAMES[19], function () { eq(r2.Categories, "Solo"); });
    check(A_NAMES[20], function () { eq(r3.Categories, ""); });
    check(A_NAMES[21], function () { eq(r1.Region, "West"); });
    check(A_NAMES[22], function () { eq(r2.Region, ""); });
    check(A_NAMES[23], function () { eq(r1.Where, "38.9, -77.03"); });
    check(A_NAMES[24], function () { eq(r3.Where, "0, 0"); });
    check(A_NAMES[25], function () {
      eq(typeof r1.Notes, "string", "Notes type");
      eq(r1.Notes, JSON.stringify({ unexpected: "shape", n: 1 }));
    });
    check(A_NAMES[26], function () {
      dataA.items.forEach(function (row, i) {
        Object.keys(row).forEach(function (k) {
          var v = row[k];
          assert(v === null || typeof v !== "object",
            "row " + i + " column '" + k + "' is a raw object: " + JSON.stringify(v));
        });
      });
    });

    check(A_NAMES[27], function () { eq(r1["Assigned To"], "Jane Doe <jane@example.gov>"); });
    check(A_NAMES[28], function () { eq(r1["Created By"], "Ada Lovelace <ada@example.gov>"); });
    check(A_NAMES[29], function () { eq(r2["Assigned To"], ""); });
    check(A_NAMES[30], function () { eq(r3["Assigned To"], "X One <x@example.gov>; Y Two <y@example.gov>"); });

    check(A_NAMES[31], function () {
      assert(expandedUrl, "no expanded /items request captured");
      contains(decodeURIComponent(expandedUrl), "CaseLink/CaseNumber", "expand-select");
      contains(decodeURIComponent(expandedUrl), "$expand=", "expand clause");
      contains(decodeURIComponent(expandedUrl), "AssignedTo/EMail", "user expansion");
    });
    check(A_NAMES[32], function () { eq(r1.Case, "WHD-2024-001"); });
    check(A_NAMES[33], function () {
      eq(typeof r2.Case, "string", "null show-value type");
      eq(r2.Case, "");
    });
    check(A_NAMES[34], function () { eq(r3.Case, "WHD-A; WHD-B"); });

    check(A_NAMES[35], function () {
      deepEq(sortedKeys(dataA), SHAPE_KEYS, "top-level key set");
      assert(relativeOrderOk(Object.keys(dataA), BASE_KEYS),
        "original keys out of order: " + JSON.stringify(Object.keys(dataA)));
      eq(dataA.web, WEB, "web (trailing slash trimmed)");
      eq(dataA.list, LIST);
      assert(!isNaN(Date.parse(dataA.exportedAt)), "exportedAt not an ISO date: " + dataA.exportedAt);
    });
    check(A_NAMES[36], function () { eq(dataA.listId, LIST_ID); });
    check(A_NAMES[37], function () {
      eq(dataA.columns.length, KEPT_FIELDS.length + 1, "column count");
      dataA.columns.forEach(function (c, i) {
        deepEq(Object.keys(c), ["name", "internal", "type"], "columns[" + i + "] keys");
        eq(typeof c.name, "string", "columns[" + i + "].name");
        eq(typeof c.internal, "string", "columns[" + i + "].internal");
        eq(typeof c.type, "string", "columns[" + i + "].type");
      });
      deepEq(dataA.columns[13], { name: "Case", internal: "CaseLink", type: "Lookup" }, "lookup column");
    });
    check(A_NAMES[38], function () {
      deepEq(dataA.columns[0], { name: "ID", internal: "Id", type: "Counter" });
    });
    check(A_NAMES[39], function () { eq(dataA.expanded, true, "expanded flag on the normal path"); });
    // The marker is the FLAT path's tell. Emitting it on a good export would
    // make every export look degraded to the MCP server and the agent.
    check(A_NAMES[40], function () {
      assert(!("degraded" in dataA),
        "a fully expanded export must not carry a degraded marker: " + JSON.stringify(dataA.degraded));
      eq(dataA.degraded, undefined, "degraded value");
    });
  } catch (e) {
    groupFailed(A_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario B: expanded query fails -> flat fallback ---------- */

  var B_NAMES = [
    "fallback: flat query issued when expanded query rejects",
    "fallback: rows still returned from the flat query",
    "fallback: User column becomes 'id:N'",
    "fallback: Lookup column becomes 'id:N'",
    "fallback: null *Id in flat mode -> ''",
    "fallback: flat select asks for AuthorId/AssignedToId/CaseLinkId",
    "fallback: warning logged about retrying flat",
    "fallback: expanded flag is false on the flat path",
    "fallback: payload carries a human-readable 'degraded' marker, not just expanded:false"
  ];

  try {
    var envB = makeEnv({ expandedFails: true });
    var splaB = loadSPLA(envB);
    var dataB = await splaB.read(LIST);
    var flatUrl = envB.requests.filter(function (u) {
      return u.indexOf("/items") !== -1 && u.indexOf("$expand=") === -1;
    })[0];

    check(B_NAMES[0], function () { assert(flatUrl, "no flat /items request was made"); });
    check(B_NAMES[1], function () {
      eq(dataB.items.length, 2, "row count");
      eq(dataB.items[0].Title, "Alpha");
    });
    check(B_NAMES[2], function () {
      eq(dataB.items[0]["Assigned To"], "id:12");
      eq(dataB.items[0]["Created By"], "id:7");
    });
    check(B_NAMES[3], function () { eq(dataB.items[0].Case, "id:3"); });
    check(B_NAMES[4], function () {
      eq(dataB.items[1]["Assigned To"], "");
      eq(dataB.items[1].Case, "");
    });
    check(B_NAMES[5], function () {
      var u = decodeURIComponent(flatUrl);
      ["AuthorId", "AssignedToId", "CaseLinkId"].forEach(function (n) { contains(u, n, "flat select"); });
      assert(u.indexOf("$expand=") === -1, "flat query must not expand");
    });
    check(B_NAMES[6], function () {
      assert(envB.warns.some(function (w) { return /retrying flat/i.test(w); }), "no flat-retry warning: " + JSON.stringify(envB.warns));
    });
    check(B_NAMES[7], function () { eq(dataB.expanded, false, "expanded flag on the flat-fallback path"); });
    // `expanded:false` is a flag a consumer has to already know the meaning of.
    // The degraded string is what tells a human (and an agent reading the JSON)
    // WHY every person/lookup cell now says "id:12" instead of a name.
    check(B_NAMES[8], function () {
      assert("degraded" in dataB, "flat-fallback export is missing the degraded marker; keys=" + JSON.stringify(Object.keys(dataB)));
      eq(typeof dataB.degraded, "string", "degraded type");
      assert(dataB.degraded.length > 0, "degraded marker is an empty string");
      assert(/id/i.test(dataB.degraded),
        "degraded marker must explain that person/lookup columns are ids: " + JSON.stringify(dataB.degraded));
      // It travels with the payload, so a written export carries it too.
      assert(JSON.parse(JSON.stringify(dataB)).degraded === dataB.degraded, "degraded must survive serialization");
    });
  } catch (e) {
    groupFailed(B_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario C: both queries fail ---------- */

  var C_NAMES = [
    "both-fail: read() rejects",
    "both-fail: message names the expanded-query failure",
    "both-fail: message names the flat-fallback failure",
    "both-fail: message names the list"
  ];

  try {
    var envC = makeEnv({ expandedFails: true, flatFails: true });
    var splaC = loadSPLA(envC);
    var errC = null;
    try {
      await splaC.read(LIST);
    } catch (e) {
      errC = e;
    }

    check(C_NAMES[0], function () { assert(errC, "read() resolved but should have thrown"); });
    check(C_NAMES[1], function () {
      contains(errC.message, "expanded query:", "label");
      contains(errC.message, EXPANDED_ERR_BODY, "expanded body");
    });
    check(C_NAMES[2], function () {
      contains(errC.message, "flat fallback:", "label");
      contains(errC.message, FLAT_ERR_BODY, "flat body");
    });
    check(C_NAMES[3], function () { contains(errC.message, "Read of '" + LIST + "' failed"); });
  } catch (e) {
    groupFailed(C_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario D: folder write via File System Access ---------- */

  var D_NAMES = [
    "folder: SPLA.folder() resolves true when the picker grants a handle",
    "folder: picker requested in readwrite mode",
    "folder: export calls getFileHandle once, with a sanitized dated filename",
    "folder: createWritable -> write -> close called in order",
    "folder: the folder holds exactly one entry, under the name getFileHandle asked for",
    "folder: that entry's COMMITTED content deep-equals the returned export",
    "folder: committed payload carries the expanded flag (true here)",
    "folder: export does NOT trigger a download",
    "folder: export still returns the data object",
    "folder: no writable stream left neither closed nor aborted",
    "folder: PERSISTENCE - a 2nd export writes into the folder again",
    "folder: PERSISTENCE - the 2nd export triggers no download",
    "folder: PERSISTENCE - the 2nd export re-commits the same file with fresh bytes"
  ];

  try {
    var envD = makeEnv({});
    var splaD = loadSPLA(envD);
    var okD = await splaD.folder();
    var dataD = await splaD.export(LIST);

    var gfhD = gfhCreates(envD);
    var seq = fsSeq(envD);

    check(D_NAMES[0], function () { eq(okD, true, "folder() return"); });
    check(D_NAMES[1], function () {
      var pick = envD.fsCalls.filter(function (c) { return c[0] === "showDirectoryPicker"; })[0];
      assert(pick, "picker never called");
      eq(pick[1], "readwrite", "picker mode");
    });
    check(D_NAMES[2], function () {
      eq(gfhD.length, 1, "creating getFileHandle call count");
      assertName(gfhD[0][1], LIST, dataD.exportedAt, "filename");
      eq(gfhD[0][2], true, "create:true option");
      // Any existence probe must be for the SAME name it then creates.
      gfhAll(envD).forEach(function (c) { eq(c[1], gfhD[0][1], "getFileHandle name"); });
    });
    check(D_NAMES[3], function () {
      deepEq(seq, ["showDirectoryPicker", "getFileHandle", "createWritable", "write", "close"], "call sequence");
      // A probe, when present, happens BEFORE the file is created.
      var all = envD.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; });
      var firstCreate = all.map(function (c) { return c[2] === true; }).indexOf(true);
      assert(all.slice(0, firstCreate).every(function (c) { return c[2] !== true; }), "probe ordering");
    });
    check(D_NAMES[4], function () {
      // The bytes must be findable under the name the folder itself stored them
      // at - not under a name recycled from the call we just recorded.
      deepEq(folderNames(envD), [gfhD[0][1]], "folder contents");
    });
    check(D_NAMES[5], function () {
      var stored = soleFolderName(envD);
      var text = committed(envD, stored);
      assert(text !== null, "entry was never committed (stream not closed)");
      assert(text.length > 0, "committed a 0-byte file");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataD)), "committed content vs returned export");
    });
    check(D_NAMES[6], function () {
      eq(JSON.parse(committed(envD, soleFolderName(envD))).expanded, true, "expanded flag in the written payload");
    });
    check(D_NAMES[7], function () {
      eq(envD.downloads.length, 0, "downloads triggered");
      eq(envD.anchors.length, 0, "anchor elements created");
      eq(envD.blobs.length, 0, "blobs created");
      eq(envD.objectUrls, 0, "object URLs created");
    });
    check(D_NAMES[8], function () {
      eq(dataD.items.length, 3);
      eq(dataD.list, LIST);
    });
    check(D_NAMES[9], function () { assertNoOrphanStreams(envD); });

    // The whole point of the feature: point it at config/exports/ ONCE and every
    // later export keeps landing there. A grant dropped after a SUCCESSFUL write
    // would send export #2 onwards to Downloads, silently, forever.
    var dataD2 = await splaD.export(LIST);

    check(D_NAMES[10], function () {
      eq(envD.writes.length, 2, "committed write count after 2 exports");
      eq(gfhCreates(envD).length, 2, "creating getFileHandle calls");
      assertName(envD.writes[1].name, LIST, dataD2.exportedAt, "2nd committed filename");
    });
    check(D_NAMES[11], function () {
      eq(envD.downloads.length, 0, "downloads after 2nd export");
      eq(envD.blobs.length, 0, "blobs after 2nd export");
      eq(envD.objectUrls, 0, "object URLs after 2nd export");
    });
    check(D_NAMES[12], function () {
      // Same list, same day -> deliberately the same artifact, refreshed.
      var stored = soleFolderName(envD);
      eq(stored, envD.writes[1].name, "2nd export wrote somewhere else");
      var text = committed(envD, stored);
      assert(text !== null, "2nd export never committed");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataD2)), "2nd committed content");
      assertNoOrphanStreams(envD);
    });
  } catch (e) {
    groupFailed(D_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario D2: filename collisions ---------- */

  var D2_NAMES = [
    "filenames: two titles sanitizing to the same stem get DIFFERENT filenames",
    "filenames: both exports survive as separate folder entries",
    "filenames: each stored file holds its own list's payload"
  ];

  try {
    var envD2 = makeEnv({});
    var splaD2 = loadSPLA(envD2);
    await splaD2.folder();
    var one = await splaD2.export(LIST);        // "My List" -> stem My_List
    var two = await splaD2.export(LIST_TWIN);   // "My@List" -> stem My_List

    var namesD2 = gfhCreates(envD2).map(function (c) { return c[1]; });

    check(D2_NAMES[0], function () {
      eq(namesD2.length, 2, "creating getFileHandle call count");
      assert(namesD2[0] !== namesD2[1],
        "distinct lists " + JSON.stringify(LIST) + " and " + JSON.stringify(LIST_TWIN) +
        " both exported as " + JSON.stringify(namesD2[0]) + " - the second silently overwrites the first");
      assertName(namesD2[0], LIST, one.exportedAt, "1st filename");
      assertName(namesD2[1], LIST_TWIN, two.exportedAt, "2nd filename");
    });
    check(D2_NAMES[1], function () {
      eq(folderNames(envD2).length, 2, "folder entry count; entries=" + JSON.stringify(folderNames(envD2)));
      eq(envD2.downloads.length, 0, "downloads");
    });
    check(D2_NAMES[2], function () {
      var titles = folderNames(envD2).map(function (n) { return JSON.parse(committed(envD2, n)).list; }).sort();
      deepEq(titles, [LIST, LIST_TWIN].slice().sort(), "list titles stored");
      eq(one.list, LIST);
      eq(two.list, LIST_TWIN);
    });
  } catch (e) {
    groupFailed(D2_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario E: folder write fails permanently -> download ---------- */

  var E_NAMES = [
    "folder-fallback: a write that keeps failing falls back to a download",
    "folder-fallback: downloaded blob is the same JSON payload",
    "folder-fallback: the failure is reported, naming the fallback",
    "folder-fallback: NO leftover entry for that filename (no 0-byte file)",
    "folder-fallback: no writable stream left neither closed nor aborted",
    "folder-fallback: the grant is KEPT - the next export re-attempts the folder",
    "folder-fallback: 2nd export downloads again while the folder is still broken",
    "folder-fallback: still no folder entries after the 2nd export",
    "folder-fallback: RECOVERY - once writes work again the export lands in the folder",
    "folder-fallback: object URL revoked after every download"
  ];

  try {
    var envE = makeEnv({ folderWriteThrows: true });
    var splaE = loadSPLA(envE);
    await splaE.folder();
    var dataE = await splaE.export(LIST);
    var fsAfterFirst = envE.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; }).length;

    check(E_NAMES[0], function () {
      eq(envE.downloads.length, 1, "download count");
      assertName(envE.downloads[0].name, LIST, dataE.exportedAt, "download filename");
      eq(envE.anchors[0].tagName, "A", "anchor tag");
      eq(envE.anchors[0].clicks, 1, "anchor clicked");
      eq(envE.anchors[0].removed, 1, "anchor removed");
      eq(envE.appended.length, 1, "anchor appended to body");
    });
    check(E_NAMES[1], function () {
      eq(envE.blobs.length, 1, "blob count");
      eq(envE.blobs[0].type, "application/json", "blob type");
      var parsed = JSON.parse(envE.blobs[0].parts[0]);
      eq(parsed.list, LIST);
      eq(parsed.items.length, 3, "downloaded row count");
      deepEq(parsed, JSON.parse(JSON.stringify(dataE)), "downloaded payload vs returned export");
    });
    check(E_NAMES[2], function () {
      assert(envE.warns.some(function (w) { return /Folder write failed/.test(w) && /falling back to download/.test(w); }),
        "no folder-write warning: " + JSON.stringify(envE.warns));
    });
    // getFileHandle(create:true) already materialized a 0-byte file. Abandoning
    // it leaves the MCP server's exports folder holding an empty JSON that looks
    // exactly like a successful export.
    check(E_NAMES[3], function () {
      deepEq(folderNames(envE), [], "folder must be empty after a failed write, found " + JSON.stringify(folderNames(envE)));
    });
    check(E_NAMES[4], function () { assertNoOrphanStreams(envE); });

    // The grant survives one bad export; the next one tries the folder again.
    await splaE.export(LIST);
    var fsAfterSecond = envE.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; }).length;

    check(E_NAMES[5], function () {
      assert(fsAfterFirst >= 1, "1st export never touched the folder");
      assert(fsAfterSecond > fsAfterFirst,
        "2nd export did not re-attempt the folder (grant was dropped after one bad write); " +
        "getFileHandle calls " + fsAfterFirst + " -> " + fsAfterSecond);
    });
    check(E_NAMES[6], function () {
      eq(envE.downloads.length, 2, "download count after 2nd export");
      assertName(envE.downloads[1].name, LIST, dataE.exportedAt, "2nd download filename");
    });
    check(E_NAMES[7], function () {
      deepEq(folderNames(envE), [], "folder entries after 2nd export");
      assertNoOrphanStreams(envE);
    });

    envE.failAllWrites = false;                       // the lock/AV scan clears
    var dataE3rd = await splaE.export(LIST);
    check(E_NAMES[8], function () {
      eq(envE.downloads.length, 2, "recovery export must NOT download");
      var stored = soleFolderName(envE);
      assertName(stored, LIST, dataE3rd.exportedAt, "recovered filename");
      deepEq(JSON.parse(committed(envE, stored)), JSON.parse(JSON.stringify(dataE3rd)), "recovered content");
      assertNoOrphanStreams(envE);
    });

    await flush(); // let the deferred URL.revokeObjectURL callbacks run
    check(E_NAMES[9], function () {
      eq(envE.objectUrls, 2, "object URLs created");
      eq(envE.revoked, envE.objectUrls, "revokeObjectURL calls vs createObjectURL calls");
      assert(envE.timers.indexOf(1000) !== -1, "revoke was not deferred by 1000ms; timers=" + JSON.stringify(envE.timers));
    });
  } catch (e) {
    groupFailed(E_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario E2: TRANSIENT write failure -> retry, stays in folder ---------- */

  var E2_NAMES = [
    "transient: a first-attempt write failure is retried, not abandoned",
    "transient: the file ends up in the FOLDER, fully committed",
    "transient: nothing was sent to Downloads",
    "transient: the folder grant survives a transient failure",
    "transient: no writable stream left neither closed nor aborted"
  ];

  try {
    var envE2 = makeEnv({ folderWriteFailTimes: 1 });
    var splaE2 = loadSPLA(envE2);
    await splaE2.folder();
    var dataE2 = await splaE2.export(LIST);

    check(E2_NAMES[0], function () {
      assert(envE2.writeAttempts >= 2,
        "only " + envE2.writeAttempts + " write attempt(s); a transient failure must be retried");
    });
    check(E2_NAMES[1], function () {
      var stored = soleFolderName(envE2);
      assertName(stored, LIST, dataE2.exportedAt, "stored filename");
      var text = committed(envE2, stored);
      assert(text !== null, "entry never committed");
      assert(text.length > 0, "committed a 0-byte file");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataE2)), "committed content");
    });
    check(E2_NAMES[2], function () {
      eq(envE2.downloads.length, 0, "downloads");
      eq(envE2.blobs.length, 0, "blobs");
    });

    // A transient failure must not revoke the grant: the next export still lands
    // in the folder rather than quietly reverting to Downloads forever.
    var dataE2b = await splaE2.export(LIST_TWIN);
    check(E2_NAMES[3], function () {
      eq(envE2.downloads.length, 0, "downloads after the follow-up export");
      eq(envE2.writes.length, 2, "committed write count");
      var nameE2b = envE2.writes[1].name;
      assert(folderNames(envE2).indexOf(nameE2b) !== -1, "follow-up export missing from the folder");
      deepEq(JSON.parse(committed(envE2, nameE2b)), JSON.parse(JSON.stringify(dataE2b)), "follow-up content");
    });
    check(E2_NAMES[4], function () { assertNoOrphanStreams(envE2); });
  } catch (e) {
    groupFailed(E2_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario E3: serialization failure is NOT a folder failure ---------- */

  var E3_NAMES = [
    "stringify-fail: export() rejects instead of silently downloading",
    "stringify-fail: nothing was downloaded",
    "stringify-fail: the folder was never touched, so no half-written entry",
    "stringify-fail: the folder grant SURVIVES (next export still writes there)",
    "stringify-fail: no writable stream left neither closed nor aborted"
  ];

  try {
    var envE3 = makeEnv({ stringifyThrows: true });
    var splaE3 = loadSPLA(envE3);
    await splaE3.folder();

    var errE3 = null;
    await checkAsync(E3_NAMES[0], async function () {
      try {
        await splaE3.export(LIST);
      } catch (e) {
        errE3 = e;
      }
      assert(errE3, "export() resolved even though the payload could not be serialized");
      contains(String(errE3 && errE3.message), "circular", "the underlying serialization error must be reported");
    });
    check(E3_NAMES[1], function () {
      eq(envE3.downloads.length, 0, "downloads");
      eq(envE3.blobs.length, 0, "blobs");
    });
    check(E3_NAMES[2], function () {
      deepEq(folderNames(envE3), [], "folder entries, found " + JSON.stringify(folderNames(envE3)));
      eq(envE3.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; }).length, 0,
        "a payload bug must not open a file handle at all");
    });

    // Un-arm the poison; the grant must still be live - a JSON.stringify bug is
    // the payload's fault, not the folder's, and must not cost the user the grant.
    envE3.stringifyThrows = false;
    var dataE3 = await splaE3.export(LIST);

    check(E3_NAMES[3], function () {
      eq(envE3.downloads.length, 0, "a serialization bug must not revoke the folder grant");
      var stored = soleFolderName(envE3);
      assertName(stored, LIST, dataE3.exportedAt, "recovered filename");
      var text = committed(envE3, stored);
      assert(text !== null, "recovery export never committed");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataE3)), "recovery content");
    });
    check(E3_NAMES[4], function () { assertNoOrphanStreams(envE3); });
  } catch (e) {
    groupFailed(E3_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario E4: a failed re-write must not destroy a good file ----------
   *
   * The distinction this scenario exists to police:
   *
   *   OVERWRITE that fails  -> the entry was already there and already good.
   *                            ws.abort() rolls the swap file back, so doing
   *                            NOTHING ELSE leaves the previous export intact.
   *                            removeEntry() here DELETES the user's last good
   *                            export because of a transient lock - strictly
   *                            worse than the 0-byte leftover it was added for.
   *   FRESH create that fails -> the entry only exists because create:true just
   *                            made it, 0 bytes. It must be removed.
   *
   * So the cleanup has to be conditional on whether the entry pre-existed, and
   * this scenario asserts BOTH halves in one folder. An earlier version of this
   * test wrapped the survival assertions in `if (entry)` and called deletion
   * "also acceptable" - that tolerance is precisely how the data-loss
   * regression got through, so the check is now unconditional.
   */

  var E4_NAMES = [
    "re-write: the pre-existing good export SURVIVES a failed overwrite (entry still exists)",
    "re-write: the survivor's committed bytes are the ORIGINAL export, byte-for-byte",
    "re-write: the payload still reaches the user via download",
    "re-write: a FRESH create that fails leaves NO entry behind",
    "re-write: no writable stream left neither closed nor aborted"
  ];

  try {
    var envE4 = makeEnv({});
    var splaE4 = loadSPLA(envE4);
    await splaE4.folder();

    // Pre-seed the folder with a known-good export and remember its exact bytes.
    var firstE4 = await splaE4.export(LIST);
    var nameE4 = soleFolderName(envE4);
    var originalBytes = committed(envE4, nameE4);
    if (originalBytes == null || originalBytes.length === 0) {
      throw new Error("scenario setup failed: the seed export did not commit any bytes");
    }

    envE4.failAllWrites = true;                 // permission/lock lost between exports
    var secondE4 = await splaE4.export(LIST);   // same list, same day -> SAME filename

    check(E4_NAMES[0], function () {
      assert(Object.prototype.hasOwnProperty.call(envE4.files, nameE4),
        "the previous good export " + JSON.stringify(nameE4) + " was DELETED by a failed overwrite - " +
        "a transient lock must not cost the user their last export; folder now holds " +
        JSON.stringify(folderNames(envE4)));
    });
    check(E4_NAMES[1], function () {
      var entry = envE4.files[nameE4];
      assert(entry, "no surviving entry to inspect (see the previous check)");
      assert(entry.committed, "the surviving entry is uncommitted - a failed write truncated a good file");
      assert(entry.content.length > 0, "the surviving entry is 0-byte - a failed write truncated a good file");
      eq(entry.content, originalBytes, "surviving bytes must be the ORIGINAL export, unmodified");
      var parsed = JSON.parse(entry.content);
      eq(parsed.list, LIST, "surviving content must be a real export");
      eq(parsed.items.length, firstE4.items.length, "surviving row count");
      deepEq(parsed, JSON.parse(JSON.stringify(firstE4)), "surviving payload vs the seed export");
    });
    check(E4_NAMES[2], function () {
      eq(envE4.downloads.length, 1, "download count");
      assertName(envE4.downloads[0].name, LIST, secondE4.exportedAt, "download filename");
      deepEq(JSON.parse(envE4.blobs[0].parts[0]), JSON.parse(JSON.stringify(secondE4)), "downloaded payload");
    });

    // Same broken folder, a filename that does NOT exist yet: here the entry is
    // purely an artifact of create:true and must not be left behind.
    var thirdE4 = await splaE4.export(LIST_TWIN);

    check(E4_NAMES[3], function () {
      // Exactly the original entry, and nothing else: proves the fresh 0-byte
      // file was cleaned up AND that the survivor is still there afterwards.
      deepEq(folderNames(envE4), [nameE4],
        "after a failed fresh create the folder must hold only the pre-existing export");
      eq(envE4.downloads.length, 2, "download count after the fresh-create failure");
      assertName(envE4.downloads[1].name, LIST_TWIN, thirdE4.exportedAt, "2nd download filename");
      eq(committed(envE4, nameE4), originalBytes, "survivor still intact after the second failure");
    });
    check(E4_NAMES[4], function () { assertNoOrphanStreams(envE4); });
  } catch (e) {
    groupFailed(E4_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario F: picker unavailable / cancelled ---------- */

  var F_NAMES = [
    "no-picker: folder() returns false and warns when the API is absent",
    "no-picker: export still downloads",
    "picker-cancelled: folder() returns false and export downloads"
  ];

  try {
    var envF = makeEnv({ noPicker: true });
    var splaF = loadSPLA(envF);
    var okF = await splaF.folder();
    await splaF.export(LIST);

    check(F_NAMES[0], function () {
      eq(okF, false, "folder() return");
      assert(envF.warns.some(function (w) { return /No File System Access API/.test(w); }), "no unsupported-browser warning");
    });
    check(F_NAMES[1], function () { eq(envF.downloads.length, 1, "download count"); });

    var envF2 = makeEnv({ pickerRejects: true });
    var splaF2 = loadSPLA(envF2);
    var okF2 = await splaF2.folder();
    await splaF2.export(LIST);
    check(F_NAMES[2], function () {
      eq(okF2, false, "folder() return on cancel");
      eq(envF2.downloads.length, 1, "download count");
      eq(envF2.fsCalls.filter(function (c) { return c[0] === "getFileHandle"; }).length, 0, "must not touch a handle");
      deepEq(folderNames(envF2), [], "must not create any folder entry");
    });
  } catch (e) {
    groupFailed(F_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario F2: a cancelled RE-pick must not destroy a live grant ----------
   *
   * Every other scenario calls folder() exactly once, so nothing tested what a
   * SECOND call does. Real usage invites one: the grant dies on page reload, so
   * users re-run `await SPLA.folder()` out of habit - and pressing Esc on that
   * dialog must not silently downgrade the rest of the page-load to Downloads.
   */

  var F2_NAMES = [
    "re-pick: a cancelled 2nd folder() reports success - the existing grant still stands",
    "re-pick: it says so, rather than reporting the folder as lost",
    "re-pick: the export AFTER the cancelled re-pick still lands in the FOLDER",
    "re-pick: nothing was sent to Downloads across the whole scenario"
  ];

  try {
    var envF3 = makeEnv({});
    var splaF3 = loadSPLA(envF3);

    var okF3a = await splaF3.folder();
    var dataF3a = await splaF3.export(LIST);
    var nameF3a = soleFolderName(envF3);

    envF3.failPicker = true;                    // user presses Esc on the re-pick
    var warnsBefore = envF3.warns.length;
    var okF3b = await splaF3.folder();

    var dataF3b = await splaF3.export(LIST_TWIN);   // different filename -> both visible

    check(F2_NAMES[0], function () {
      eq(okF3a, true, "1st folder() return");
      eq(envF3.pickerCalls, 2, "picker call count");
      eq(okF3b, true,
        "folder() reported failure after a cancelled re-pick even though a working grant was already held - " +
        "callers read that as 'exports go to Downloads now'");
    });
    check(F2_NAMES[1], function () {
      var fresh = envF3.warns.slice(warnsBefore);
      assert(fresh.some(function (w) { return /KEEPING|still write|already granted/i.test(w); }),
        "no message that the existing folder was kept: " + JSON.stringify(fresh));
    });
    check(F2_NAMES[2], function () {
      var names = folderNames(envF3);
      eq(names.length, 2, "folder entry count after both exports; entries=" + JSON.stringify(names));
      assert(names.indexOf(nameF3a) !== -1, "the first export vanished");
      var second = names.filter(function (n) { return n !== nameF3a; })[0];
      assertName(second, LIST_TWIN, dataF3b.exportedAt, "2nd export filename");
      var text = committed(envF3, second);
      assert(text !== null, "2nd export never committed");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataF3b)), "2nd export content");
      deepEq(JSON.parse(committed(envF3, nameF3a)), JSON.parse(JSON.stringify(dataF3a)), "1st export content");
    });
    check(F2_NAMES[3], function () {
      eq(envF3.downloads.length, 0, "downloads");
      eq(envF3.blobs.length, 0, "blobs");
      assertNoOrphanStreams(envF3);
    });
  } catch (e) {
    groupFailed(F2_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario P: the permission re-check after a failed write ----------
   *
   * A folder write can fail for two very different reasons that look identical
   * at the call site: a transient lock (VS Code, OneDrive, AV) or a grant that
   * is actually gone. The reader keeps the handle either way and re-verifies via
   * queryPermission/requestPermission on the NEXT export. Verified here:
   *   granted  -> keep writing to the folder
   *   prompt   -> escalate to requestPermission
   *   denied   -> drop the handle, tell the user how to re-grant, download
   *   absent   -> old engine, treat as usable (never refuse over a missing API)
   *   throws   -> un-checkable, attempt the write anyway
   */

  var P_NAMES = [
    "permission: an untroubled export does NOT consult queryPermission",
    "permission: after a failed write the next export DOES consult queryPermission (readwrite)",
    "permission: 'granted' keeps the folder - the export lands there, no download",
    "permission: the re-check is one-shot; a clean export afterwards does not re-query",
    "permission: 'prompt' escalates to requestPermission, and 'granted' there keeps the folder",
    "permission: 'denied' consults both query and request",
    "permission: 'denied' DROPS the handle - the folder is not touched at all",
    "permission: 'denied' downloads instead, and says how to re-grant the folder",
    "permission: once dropped the grant stays dropped - later exports download without re-querying",
    "permission: an engine WITHOUT queryPermission/requestPermission is treated as usable",
    "permission: a queryPermission that throws is not treated as a denial"
  ];

  try {
    /* -- granted -- */
    var envP1 = makeEnv({ folderWriteThrows: true, permission: { query: "granted" } });
    var splaP1 = loadSPLA(envP1);
    await splaP1.folder();
    await splaP1.export(LIST);                  // fails twice -> download, arms the re-check
    var permAfterFirstP1 = envP1.permCalls.length;

    envP1.failAllWrites = false;                // the lock clears; the grant was fine all along
    var dataP1 = await splaP1.export(LIST);

    check(P_NAMES[0], function () {
      eq(permAfterFirstP1, 0,
        "the first export re-checked a permission that had never failed: " + JSON.stringify(envP1.permCalls));
      eq(envP1.downloads.length, 1, "first export should have fallen back to download");
    });
    check(P_NAMES[1], function () {
      assert(envP1.permCalls.length > 0,
        "the export after a failed write never consulted queryPermission - a revoked grant would go unnoticed");
      eq(envP1.permCalls[0][0], "queryPermission", "first permission call");
      eq(envP1.permCalls[0][1], "readwrite", "permission mode asked for");
    });
    check(P_NAMES[2], function () {
      eq(envP1.downloads.length, 1, "a granted re-check must not download");
      var stored = soleFolderName(envP1);
      assertName(stored, LIST, dataP1.exportedAt, "stored filename");
      var text = committed(envP1, stored);
      assert(text !== null, "entry never committed");
      deepEq(JSON.parse(text), JSON.parse(JSON.stringify(dataP1)), "committed content");
      assertNoOrphanStreams(envP1);
    });

    var permAfterRecoveryP1 = envP1.permCalls.length;
    await splaP1.export(LIST);
    check(P_NAMES[3], function () {
      eq(envP1.permCalls.length, permAfterRecoveryP1,
        "a healthy export re-queried the permission; the flag must be consumed by the re-check that used it");
      eq(envP1.downloads.length, 1, "downloads after the third export");
    });

    /* -- prompt -> request -- */
    var envP2 = makeEnv({ folderWriteThrows: true, permission: { query: "prompt", request: "granted" } });
    var splaP2 = loadSPLA(envP2);
    await splaP2.folder();
    await splaP2.export(LIST);
    envP2.failAllWrites = false;
    var dataP2 = await splaP2.export(LIST);

    check(P_NAMES[4], function () {
      deepEq(permNames(envP2), ["queryPermission", "requestPermission"], "permission calls");
      eq(envP2.permCalls[1][1], "readwrite", "requestPermission mode");
      eq(envP2.downloads.length, 1, "a re-granted permission must not download");
      var stored = soleFolderName(envP2);
      deepEq(JSON.parse(committed(envP2, stored)), JSON.parse(JSON.stringify(dataP2)), "committed content");
    });

    /* -- denied -- */
    var envP3 = makeEnv({ folderWriteThrows: true, permission: { query: "denied", request: "denied" } });
    var splaP3 = loadSPLA(envP3);
    await splaP3.folder();
    await splaP3.export(LIST);                  // download #1, arms the re-check

    // Writes would now SUCCEED. Only an honest permission re-check keeps this
    // export out of the folder; skipping the check silently "works", which is
    // exactly the bug - the handle is stale and the user is never told.
    envP3.failAllWrites = false;
    var gfhBeforeP3 = gfhAll(envP3).length;
    var warnsBeforeP3 = envP3.warns.length;
    var dataP3 = await splaP3.export(LIST);

    check(P_NAMES[5], function () {
      deepEq(permNames(envP3), ["queryPermission", "requestPermission"], "permission calls");
    });
    check(P_NAMES[6], function () {
      eq(gfhAll(envP3).length, gfhBeforeP3,
        "the export touched the folder after the permission came back denied");
      deepEq(folderNames(envP3), [], "folder entries; a denied grant must leave the folder alone");
      assertNoOrphanStreams(envP3);
    });
    check(P_NAMES[7], function () {
      eq(envP3.downloads.length, 2, "download count");
      assertName(envP3.downloads[1].name, LIST, dataP3.exportedAt, "download filename");
      var fresh = envP3.warns.slice(warnsBeforeP3);
      var told = fresh.filter(function (w) {
        return w.indexOf("ERROR") === 0 && /permission/i.test(w) && /SPLA\.folder\(\)/.test(w);
      });
      assert(told.length > 0,
        "no console.error telling the user the grant is gone and to re-run `await SPLA.folder()`: " + JSON.stringify(fresh));
    });

    var permAfterDenyP3 = envP3.permCalls.length;
    var dataP3b = await splaP3.export(LIST);
    check(P_NAMES[8], function () {
      eq(envP3.permCalls.length, permAfterDenyP3, "re-queried a handle that was already dropped");
      eq(envP3.downloads.length, 3, "download count");
      assertName(envP3.downloads[2].name, LIST, dataP3b.exportedAt, "download filename");
      deepEq(folderNames(envP3), [], "folder entries");
    });

    /* -- API absent (older engine) -- */
    var envP4 = makeEnv({ folderWriteThrows: true });   // no `permission` -> no methods at all
    var splaP4 = loadSPLA(envP4);
    await splaP4.folder();
    await splaP4.export(LIST);
    envP4.failAllWrites = false;
    var dataP4 = await splaP4.export(LIST);

    check(P_NAMES[9], function () {
      eq(envP4.perm, null, "scenario models an engine with no permission API");
      deepEq(permNames(envP4), [], "no permission calls are possible");
      eq(envP4.downloads.length, 1, "an unverifiable grant must not be treated as denied");
      var stored = soleFolderName(envP4);
      deepEq(JSON.parse(committed(envP4, stored)), JSON.parse(JSON.stringify(dataP4)), "committed content");
      assertNoOrphanStreams(envP4);
    });

    /* -- query throws -- */
    var envP5 = makeEnv({ folderWriteThrows: true, permission: { query: "granted", queryThrows: true } });
    var splaP5 = loadSPLA(envP5);
    await splaP5.folder();
    await splaP5.export(LIST);
    envP5.failAllWrites = false;
    var dataP5 = await splaP5.export(LIST);

    check(P_NAMES[10], function () {
      deepEq(permNames(envP5), ["queryPermission"], "permission calls");
      eq(envP5.downloads.length, 1, "a failed permission PROBE must not cost the user the folder");
      var stored = soleFolderName(envP5);
      deepEq(JSON.parse(committed(envP5, stored)), JSON.parse(JSON.stringify(dataP5)), "committed content");
      assertNoOrphanStreams(envP5);
    });
  } catch (e) {
    groupFailed(P_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario G: context + lists (supporting behaviour) ---------- */

  var G_NAMES = [
    "context: web() trims the trailing slash off webAbsoluteUrl",
    "context: web() falls back to location.origin and warns without _spPageContextInfo",
    "lists: non-hidden lists returned, sorted by Title",
    "list title with an apostrophe is OData-escaped"
  ];

  try {
    var envG = makeEnv({});
    var splaG = loadSPLA(envG);
    check(G_NAMES[0], function () { eq(splaG.web(), WEB); });

    var envG2 = makeEnv({ noContext: true });
    var splaG2 = loadSPLA(envG2);
    check(G_NAMES[1], function () {
      eq(splaG2.web(), "https://contoso.sharepoint.com");
      assert(envG2.warns.some(function (w) { return /No _spPageContextInfo/.test(w); }), "no context warning");
    });

    var rowsG = await splaG.lists();
    check(G_NAMES[2], function () {
      deepEq(rowsG.map(function (r) { return r.Title; }), ["My List", "Zeta Library"], "sorted titles");
      deepEq(rowsG[0], { Title: "My List", Items: 3, Template: 100 }, "row shape");
    });

    var envG3 = makeEnv({});
    var splaG3 = loadSPLA(envG3);
    try { await splaG3.read("Bob's List"); } catch (e) { /* mock has no route; the URL is what matters */ }
    check(G_NAMES[3], function () {
      var u = envG3.requests[0] || "";
      contains(u, "getByTitle('Bob''s%20List')", "apostrophe doubled and encoded");
    });
  } catch (e) {
    groupFailed(G_NAMES, e && e.message ? e.message : String(e));
  }

  /* ---------- Scenario H: the harness itself ---------- */

  var H_NAMES = [
    "harness: check() FAILS an assertion that returns a promise",
    "harness: checkAsync() FAILS a rejecting async assertion"
  ];

  // These run against a scratch results array so the meta-tests do not pollute
  // the real report - they assert that a bad assertion is caught, not skipped.
  var realResults = results;
  results = [];
  check("probe: promise-returning assertion", async function () { throw new Error("this list is empty!"); });
  await checkAsync("probe: rejecting async assertion", async function () { throw new Error("boom"); });
  var probes = results;
  results = realResults;

  check(H_NAMES[0], function () {
    assert(probes[0] && probes[0].ok === false, "check() recorded PASS for a promise-returning assertion");
    contains(probes[0].err, "returned a promise", "error message");
  });
  check(H_NAMES[1], function () {
    assert(probes[1] && probes[1].ok === false, "checkAsync() recorded PASS for a rejecting assertion");
    contains(probes[1].err, "boom", "error message");
  });

  /* ------------------------------------------------------------------ *
   * Report
   * ------------------------------------------------------------------ */

  var passed = 0, failed = 0;
  results.forEach(function (r) {
    if (r.ok) { passed++; console.log("PASS  " + r.name); }
    else { failed++; console.log("FAIL  " + r.name + "\n        " + r.err); }
  });

  console.log("\n" + passed + "/" + results.length + " checks passed" + (failed ? ", " + failed + " FAILED" : ""));
  process.exitCode = failed ? 1 : 0;
}

main().catch(function (e) {
  console.error("harness crashed: " + (e && e.stack ? e.stack : e));
  process.exitCode = 1;
});
