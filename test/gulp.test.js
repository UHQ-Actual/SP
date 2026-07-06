/*
 * End-to-end runtime test (no live network). Boots the snippet + standalone HTML
 * in jsdom with mocked SharePoint REST responses and drives the Gulp flow.
 *   cd test && npm install jsdom && node gulp.test.js
 */
const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");
const SNIPPET = fs.readFileSync(path.join(__dirname, "..", "sharepoint-console-plus.js"), "utf8");
const HTML = fs.readFileSync(path.join(__dirname, "..", "sharepoint-console.html"), "utf8");

const LISTS = { value: [
  { Title:"Enforcement Cases", Id:"11111111-1111-1111-1111-111111111111", BaseType:0, BaseTemplate:100, ItemCount:42, Hidden:false, DefaultViewUrl:"/sites/whd/mw/mwplanning/Lists/EnforcementCases/AllItems.aspx", LastItemModifiedDate:"2026-07-01T12:00:00Z" },
  { Title:"Shared Documents",  Id:"22222222-2222-2222-2222-222222222222", BaseType:1, BaseTemplate:101, ItemCount:7,  Hidden:false, DefaultViewUrl:"/sites/whd/mw/mwplanning/Shared Documents/Forms/AllItems.aspx", LastItemModifiedDate:"2026-06-15T09:30:00Z" },
  { Title:"appdata",           Id:"33333333-3333-3333-3333-333333333333", BaseType:1, BaseTemplate:101, ItemCount:3,  Hidden:true,  DefaultViewUrl:"/x", LastItemModifiedDate:"2026-01-01T00:00:00Z" },
  { Title:"Composed Looks",    Id:"44444444-4444-4444-4444-444444444444", BaseType:0, BaseTemplate:124, ItemCount:9,  Hidden:false, DefaultViewUrl:"/y", LastItemModifiedDate:"2026-01-01T00:00:00Z" }
]};
const ITEMS = { value: [
  { "odata.type":"x", Id:1, Title:"Case A", Status:"Open",   Region:"West", Assigned:{ Title:"J. Uphold" } },
  { "odata.type":"x", Id:2, Title:"Case B", Status:"Closed", Region:"East", Assigned:{ Title:"K. Doe" } }
]};
let lastBlob = null;
const wait = ms => new Promise(r=>setTimeout(r,ms));
const stubFetch = (url) => {
  let data = url.indexOf("/items")>=0 ? ITEMS : (url.indexOf("/_api/web/lists")>=0 ? LISTS : {});
  return Promise.resolve({ ok:true, status:200, json:()=>Promise.resolve(data) });
};

(async () => {
  let failures = 0;
  const ok = (c,m) => { console.log((c?"  PASS ":"  FAIL ")+m); if(!c) failures++; };

  const dom = new JSDOM("<!doctype html><html><body></body></html>", {
    url: "https://usdol.sharepoint.com/sites/whd/mw/mwplanning/Lists/EnforcementCases/AllItems.aspx",
    runScripts: "dangerously", pretendToBeVisual: true,
    beforeParse(win){
      win._spPageContextInfo = { webAbsoluteUrl:"https://usdol.sharepoint.com/sites/whd/mw/mwplanning", webServerRelativeUrl:"/sites/whd/mw/mwplanning", pageListId:"{11111111-1111-1111-1111-111111111111}" };
      win.fetch = stubFetch;
      win.URL.createObjectURL = (b)=>{ lastBlob=b; return "blob:mock"; };
      win.URL.revokeObjectURL = ()=>{};
    }
  });
  const win = dom.window;
  const s = win.document.createElement("script"); s.textContent = SNIPPET; win.document.body.appendChild(s);

  const hostEl = win.document.getElementById("spcp-host");
  ok(!!hostEl, "panel host appended to document");
  const root = hostEl && hostEl.shadowRoot;
  ok(!!root, "panel mounted in shadow DOM");
  ok(!!win.__SPCP, "window.__SPCP toggle registered");

  const gulpTab = [...root.querySelectorAll(".tab")].find(t=>t.textContent==="Gulp");
  ok(!!gulpTab && !gulpTab.classList.contains("disabled"), "Gulp tab present + enabled");
  gulpTab.click();
  const baseInput = root.querySelector(".body input.search");
  ok(!!baseInput && baseInput.value.indexOf("mwplanning")>=0, "base URL pre-seeded: "+(baseInput&&baseInput.value));

  const gulpBtn = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="Gulp lists");
  ok(!!gulpBtn, "Gulp lists button present"); gulpBtn.click(); await wait(60);

  const headerRow = root.querySelector(".tbl thead tr");
  const bodyRows = root.querySelectorAll(".tbl tbody tr");
  ok(!!headerRow, "lists table rendered");
  ok(bodyRows.length===2, "hidden + catalog filtered → 2 content lists (got "+bodyRows.length+")");
  ok(headerRow&&[...headerRow.cells].map(c=>c.textContent).join(",")==="Title,Type,Items,Modified,Url", "columns correct");
  const titles = [...bodyRows].map(r=>r.cells[0].textContent).sort().join("|");
  ok(titles==="Enforcement Cases|Shared Documents", "custom list (BaseTemplate 100) kept: "+titles);
  ok(!!root.querySelector(".tbl a.tlink"), "URL cell rendered as link");

  const stored = JSON.parse(win.localStorage.getItem("sp_link_console_v1")||"[]");
  ok(stored.length===2, "2 URLs persisted to JSON store (got "+stored.length+")");
  ok(stored.every(x=>x.url&&x.type), "stored links parsed (url+type)");

  root.querySelector(".tbl tr.clickable").click(); await wait(60);
  const itemRows = root.querySelectorAll(".tbl tbody tr");
  ok(itemRows.length===2, "items table 2 rows (got "+itemRows.length+")");
  const itemHead = [...root.querySelector(".tbl thead tr").cells].map(c=>c.textContent);
  ok(itemHead.includes("Status")&&itemHead.includes("Region"), "item columns from data: "+itemHead.join(","));
  const ai = itemHead.indexOf("Assigned");
  ok(ai>=0 && root.querySelector(".tbl tbody tr").cells[ai].textContent==="J. Uphold", "person/lookup flattened to Title");

  const mdBtn = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="Save .md");
  ok(mdBtn&&!mdBtn.classList.contains("is-disabled"), "Save .md enabled after gulp");
  lastBlob=null; mdBtn.click(); await wait(20);
  ok(!!lastBlob, "Save .md produced a Blob");
  if(lastBlob){ const md=await lastBlob.text();
    ok(/\|\s*Id\s*\|/.test(md), "markdown has header row");
    ok(/\|\s*---\s*\|/.test(md), "markdown has separator row (VS Code preview)"); }

  // ---- "This open list" (uses page context, no crawl needed) ----
  gulpTab.click(); // re-render tab fresh
  const openBtn = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="This open list");
  ok(!!openBtn, "'This open list' button present");
  ok(!openBtn.classList.contains("is-disabled"), "'This open list' enabled when a list is open (ctx.listId)");
  openBtn.click(); await wait(60);
  ok(root.querySelectorAll(".tbl tbody tr").length===2, "open-list gulp rendered its items (got "+root.querySelectorAll(".tbl tbody tr").length+")");

  // ---- "All lists' items" bulk pull → sectioned bundle ----
  const allBtn = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="All lists' items");
  ok(!!allBtn, "'All lists' items' button present");
  allBtn.click(); await wait(120);
  const sumRows = root.querySelectorAll(".tbl tbody tr");
  ok(sumRows.length===2, "bulk pull summarized 2 content lists (got "+sumRows.length+")");
  const sumHead = [...root.querySelector(".tbl thead tr").cells].map(c=>c.textContent).join(",");
  ok(sumHead==="List,Type,Items", "bulk summary columns: "+sumHead);
  // CSV disabled for the multi-list bundle; md/json/html enabled
  const csvBtn2 = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="Save .csv");
  const jsonBtn2 = [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="Save .json");
  ok(csvBtn2.classList.contains("is-disabled"), "CSV disabled for multi-list bundle");
  ok(!jsonBtn2.classList.contains("is-disabled"), "JSON enabled for multi-list bundle");
  // JSON bundle is keyed by list title, each an array of item rows
  lastBlob=null; jsonBtn2.click(); await wait(20);
  ok(!!lastBlob, "bundle .json produced a Blob");
  if(lastBlob){ const obj=JSON.parse(await lastBlob.text());
    const keys=Object.keys(obj);
    ok(keys.length===2 && Array.isArray(obj[keys[0]]), "bundle JSON keyed by list → item arrays: "+keys.join("|"));
    ok(obj[keys[0]].length===2, "each list section carries its items"); }
  // MD bundle has one section (## heading) per list
  lastBlob=null; [...root.querySelectorAll(".body .btn")].find(b=>b.textContent==="Save .md").click(); await wait(20);
  if(lastBlob){ const md=await lastBlob.text();
    ok((md.match(/^## /gm)||[]).length===2, "bundle .md has a section per list"); }

  win.__SPCP.toggle(); ok(hostEl.style.display==="none", "toggle hides panel");
  win.__SPCP.toggle(); ok(hostEl.style.display==="block", "toggle re-shows panel");

  // standalone HTML
  const dom2 = new JSDOM(HTML, { url:"https://localhost/x.html", runScripts:"dangerously", pretendToBeVisual:true });
  await wait(40);
  const d2 = dom2.window.document;
  d2.getElementById("sample").click(); await wait(20);
  ok(d2.querySelectorAll("#board .card").length===6, "standalone: 6 sample cards (got "+d2.querySelectorAll("#board .card").length+")");
  ok(d2.querySelectorAll("#chips .chip").length>=4, "standalone: filter chips rendered");
  d2.getElementById("themeToggle").click();
  ok(d2.body.classList.contains("tc-light"), "standalone: theme toggle → light");

  console.log("\n"+(failures?("FAILURES: "+failures):"ALL CHECKS PASSED"));
  process.exit(failures?1:0);
})().catch(e=>{ console.log("HARNESS ERROR:", e.stack); process.exit(2); });
