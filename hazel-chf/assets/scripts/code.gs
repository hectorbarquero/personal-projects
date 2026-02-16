/** @OnlyCurrentDoc */

/**
 * Hazel SRR + Costs
 * - Sidebar/dialog UI (uses Sidebar.html)
 * - Export: creates hazel-report.html in Google Drive for download/hosting
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Hazel SRR")
    .addItem("Open SRR sidebar", "showSrrSidebar")
    .addItem("Open Hazel dialog (wide)", "showHazelDialog")
    .addSeparator()
    .addItem("Export HTML to Drive (download after)", "exportHazelHtmlToDrive")
    .addToUi();
}

function showSrrSidebar() {
  const html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("Hazel report");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showHazelDialog() {
  const html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setWidth(1200)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, "Hazel report");
}

/** For Sidebar.html (reads active sheet, as you already have) */
function getSrrRows() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h || "").trim());
  const idx = {
    date: headers.findIndex(h => /^date$/i.test(h)),
    srr1: headers.findIndex(h => /^srr1$/i.test(h)),
    srr2: headers.findIndex(h => /^srr2$/i.test(h)),
    srr3: headers.findIndex(h => /^srr3$/i.test(h)),
    syncope: headers.findIndex(h => /^syncope$/i.test(h)),
    notes: headers.findIndex(h => /^notes$/i.test(h)),
  };

  if (idx.date < 0 || idx.srr1 < 0 || idx.srr2 < 0 || idx.srr3 < 0) {
    throw new Error("Missing required headers. Need: Date, SRR1, SRR2, SRR3 (Syncope/Notes optional).");
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateStr = (row[idx.date] || "").trim();
    if (!dateStr) continue;

    out.push({
      date: dateStr,
      srr1: toNumOrNull_(row[idx.srr1]),
      srr2: toNumOrNull_(row[idx.srr2]),
      srr3: toNumOrNull_(row[idx.srr3]),
      syncope: idx.syncope >= 0 ? toNumOrNull_(row[idx.syncope]) : null,
      notes: idx.notes >= 0 ? String(row[idx.notes] || "").trim() : "",
    });
  }
  return out;
}

function getCostsRows() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Healthcare costs");
  if (!sh) throw new Error('Sheet not found: "Healthcare costs"');

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h || "").trim());
  const idx = {
    date: headers.findIndex(h => /^date$/i.test(h)),
    item: headers.findIndex(h => /^item$/i.test(h)),
    notes: headers.findIndex(h => /^notes$/i.test(h)),
    cost: headers.findIndex(h => /^(cost|costs)$/i.test(h)),
  };
  if (idx.date < 0 || idx.item < 0 || idx.cost < 0) {
    throw new Error('Missing required headers on "Healthcare costs": Date, Item, Costs');
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateStr = (row[idx.date] || "").trim();
    if (!dateStr) continue;

    out.push({
      date: dateStr,
      item: String(row[idx.item] || "").trim(),
      notes: idx.notes >= 0 ? String(row[idx.notes] || "").trim() : "",
      cost: toMoney_(row[idx.cost]),
    });
  }
  return out;
}

function toNumOrNull_(v) {
  const s = String(v ?? "").trim();
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function toMoney_(v) {
  const s = String(v ?? "").trim();
  if (!s) return 0;
  const n = Number(s.replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

/**
 * EXPORT: Creates a standalone hazel-report.html in Google Drive.
 * Then you manually download it (Drive file menu -> Download), which saves to Downloads.
 */
function exportHazelHtmlToDrive() {
  const ss = SpreadsheetApp.getActive();

  // pinned to specific tabs so export is consistent
  const srr = getSrrRowsFromSheet_(ss, "SRR tracking");
  const costs = getCostsRowsFromSheet_(ss, "Healthcare costs");

  const html = buildStandaloneHtml_(ss.getName(), srr, costs);

  const fileName = "hazel-report.html";
  // optional: delete existing file with same name in root (comment out if you want versions)
  // deleteNamedHtmlInRoot_(fileName);

  const file = DriveApp.createFile(fileName, html, MimeType.HTML);

  SpreadsheetApp.getUi().alert(
    "Export complete",
    'Created "hazel-report.html" in Google Drive.\n\nTo put it in Downloads: open Drive, right-click the file, Download.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log("Drive file URL: " + file.getUrl());
}

function getSrrRowsFromSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h || "").trim());
  const idx = {
    date: headers.findIndex(h => /^date$/i.test(h)),
    srr1: headers.findIndex(h => /^srr1$/i.test(h)),
    srr2: headers.findIndex(h => /^srr2$/i.test(h)),
    srr3: headers.findIndex(h => /^srr3$/i.test(h)),
    syncope: headers.findIndex(h => /^syncope$/i.test(h)),
    notes: headers.findIndex(h => /^notes$/i.test(h)),
  };
  if (idx.date < 0 || idx.srr1 < 0 || idx.srr2 < 0 || idx.srr3 < 0) {
    throw new Error(`Missing required headers on "${sheetName}"`);
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateStr = (row[idx.date] || "").trim();
    if (!dateStr) continue;

    out.push({
      date: dateStr,
      srr1: toNumOrNull_(row[idx.srr1]),
      srr2: toNumOrNull_(row[idx.srr2]),
      srr3: toNumOrNull_(row[idx.srr3]),
      syncope: idx.syncope >= 0 ? toNumOrNull_(row[idx.syncope]) : null,
      notes: idx.notes >= 0 ? String(row[idx.notes] || "").trim() : "",
    });
  }
  return out;
}

function getCostsRowsFromSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => (h || "").trim());
  const idx = {
    date: headers.findIndex(h => /^date$/i.test(h)),
    item: headers.findIndex(h => /^item$/i.test(h)),
    notes: headers.findIndex(h => /^notes$/i.test(h)),
    cost: headers.findIndex(h => /^(cost|costs)$/i.test(h)),
  };
  if (idx.date < 0 || idx.item < 0 || idx.cost < 0) {
    throw new Error(`Missing required headers on "${sheetName}"`);
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dateStr = (row[idx.date] || "").trim();
    if (!dateStr) continue;

    out.push({
      date: dateStr,
      item: String(row[idx.item] || "").trim(),
      notes: idx.notes >= 0 ? String(row[idx.notes] || "").trim() : "",
      cost: toMoney_(row[idx.cost]),
    });
  }
  return out;
}

function buildStandaloneHtml_(title, srrRows, costRows) {
  const payload = {
    title,
    generatedAt: new Date().toISOString(),
    srr: srrRows,
    costs: costRows
  };

  // single static HTML file for GH pages.
  // It loads Chart.js from jsdelivr CDN.
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>${escapeHtml_(title)} - Hazel report</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <style>
    body { font-family: system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; margin: 24px; }
    h1 { margin: 0 0 6px 0; font-size: 18px; }
    h2 { margin: 22px 0 8px 0; font-size: 14px; }
    .muted { color:#666; font-size:12px; }
    .canvasWrap { height: 320px; margin: 14px 0; }
    table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 12px; }
    th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
    th { background: #f6f6f6; text-align: left; }
    .num { text-align: right; white-space: nowrap; }
    .note { white-space: pre-wrap; }
    .heatWrap { display:grid; grid-template-columns: repeat(53, 10px); gap: 3px; margin-top: 10px; }
    .heatCell { width:10px; height:10px; border-radius:2px; background:#eee; }
  </style>
</head>
<body>
  <h1>${escapeHtml_(title)} - Hazel report</h1>
  <div class="muted">Generated: ${escapeHtml_(payload.generatedAt)}</div>

  <h2>SRR</h2>
  <div class="canvasWrap"><canvas id="srrChart"></canvas></div>
  <table>
    <thead><tr>
      <th>Date</th><th>SRR1</th><th>SRR2</th><th>SRR3</th><th>Syncope</th><th>Notes</th>
    </tr></thead>
    <tbody id="srrBody"></tbody>
  </table>

  <h2>Costs</h2>
  <div id="heatmap"></div>
  <div class="canvasWrap"><canvas id="costBar"></canvas></div>
  <div class="canvasWrap"><canvas id="costLine"></canvas></div>
  <div class="canvasWrap"><canvas id="costPie"></canvas></div>

<script>
  const DATA = ${JSON.stringify(payload)};

  const esc = (s) => String(s ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
    .replaceAll('"',"&quot;").replaceAll("'","&#039;");

  function parseDateLoose(s) {
    const d = new Date(s);
    if (!isNaN(d)) return d;
    const m = String(s).match(/^(\\d{1,2})\\s+([A-Za-z]{3})\\s+(\\d{4})$/);
    if (m) return new Date(\`\${m[2]} \${m[1]} \${m[3]}\`);
    return new Date(NaN);
  }

  function bucketForItem(item) {
    const s = (item||"").toLowerCase();
    if (s.includes("imaging") || s.includes("diagnostic")) return "Imaging/Diagnostics";
    if (s.includes("emergency")) return "Emergency";
    if (s.includes("vetmedin")) return "Vetmedin";
    if (s.includes("furosemide")) return "Furosemide";
    if (s.includes("cardalis")) return "Cardalis";
    if (s.includes("hydrocodone")) return "Hydrocodone";
    if (s.includes("proviable")) return "Proviable";
    return "Other";
  }

  // SRR
  const srr = DATA.srr || [];
  new Chart(document.getElementById("srrChart"), {
    type:"line",
    data:{
      labels: srr.map(r => r.date),
      datasets:[
        { label:"SRR1", data: srr.map(r=>r.srr1), borderWidth:1.5, pointRadius:0, spanGaps:true },
        { label:"SRR2", data: srr.map(r=>r.srr2), borderWidth:1.5, pointRadius:0, spanGaps:true },
        { label:"SRR3", data: srr.map(r=>r.srr3), borderWidth:1.5, pointRadius:0, spanGaps:true }
      ]
    },
    options:{ responsive:true, maintainAspectRatio:false, interaction:{mode:"index",intersect:false} }
  });

  document.getElementById("srrBody").innerHTML = srr.map(r => \`
    <tr>
      <td>\${esc(r.date)}</td>
      <td class="num">\${r.srr1 ?? ""}</td>
      <td class="num">\${r.srr2 ?? ""}</td>
      <td class="num">\${r.srr3 ?? ""}</td>
      <td class="num">\${r.syncope ?? ""}</td>
      <td class="note">\${esc(r.notes)}</td>
    </tr>
  \`).join("");

  // Costs heatmap ... not sure how i can improve this yet
  const costs = DATA.costs || [];
  const daily = new Map();
  for (const r of costs) {
    const d = parseDateLoose(r.date);
    if (isNaN(d)) continue;
    const k = d.toISOString().slice(0,10);
    daily.set(k, (daily.get(k)||0) + (r.cost||0));
  }
  function heatColor(level) { return ["#eee","#c6dbef","#9ecae1","#6baed6","#2171b5"][level]; }

  const keys = [...daily.keys()].sort();
  if (keys.length) {
    const start = new Date(keys[0]);
    start.setDate(start.getDate() - start.getDay());
    const end = new Date(keys[keys.length-1]);
    const days = [];
    for (let d=new Date(start); d<=end; d.setDate(d.getDate()+1)) {
      const k = d.toISOString().slice(0,10);
      days.push({ k, v: daily.get(k)||0 });
    }
    const max = Math.max(...days.map(x=>x.v));
    const wrap = document.createElement("div");
    wrap.className = "heatWrap";
    for (const {k,v} of days) {
      const lvl = v<=0 ? 0 : Math.min(4, Math.ceil((v/max)*4));
      const cell = document.createElement("div");
      cell.className = "heatCell";
      cell.style.background = heatColor(lvl);
      cell.title = \`\${k}: $\${v.toFixed(2)}\`;
      wrap.appendChild(cell);
    }
    document.getElementById("heatmap").appendChild(wrap);
  }

  // Costs aggregation
  const byBucket = {};
  for (const r of costs) {
    const b = bucketForItem(r.item);
    byBucket[b] = (byBucket[b]||0) + (r.cost||0);
  }
  const bLabels = Object.keys(byBucket);
  const bValues = bLabels.map(k => byBucket[k]);

  new Chart(document.getElementById("costBar"), { type:"bar", data:{ labels:bLabels, datasets:[{ data:bValues, label:"Total spend" }] }});
  new Chart(document.getElementById("costPie"), { type:"pie", data:{ labels:bLabels, datasets:[{ data:bValues }] }});

  // Line (entry + cumulative)
  const sorted = [...costs].sort((a,b)=>parseDateLoose(a.date)-parseDateLoose(b.date));
  let run = 0;
  const cumulative = sorted.map(r => (run += (r.cost||0)));
  new Chart(document.getElementById("costLine"), {
    type:"line",
    data:{
      labels: sorted.map(r=>r.date),
      datasets:[
        { label:"Entry", data: sorted.map(r=>r.cost||0), borderWidth:1.5, pointRadius:2 },
        { label:"Cumulative", data: cumulative, borderWidth:1.5, pointRadius:0 }
      ]
    }
  });
</script>
</body>
</html>`;
}

function escapeHtml_(s) {
  return String(s ?? "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");
}
