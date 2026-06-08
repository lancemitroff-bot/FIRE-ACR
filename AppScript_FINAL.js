// ================================================================
//  FIRE DEPARTMENT ACR — Google Apps Script FINAL
//  Delete everything, paste this, re-deploy as new version
// ================================================================

const CALLS_SHEET  = "ACR Submissions";
const CONFIG_SHEET = "Config";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "saveConfig") {
      saveConfig(data.roster || [], data.units || []);
      return respond({ success: true });
    }

    // Save call report
    const sheet      = getCallsSheet();
    const asgn       = data.assignments || {};
    const byUnit     = Object.entries(asgn).map(([u,p])=>`${u}: ${p.join(", ")}`).join(" | ");
    const total      = [...new Set(Object.values(asgn).flat())].length;

    sheet.appendRow([
      new Date().toLocaleString("en-US"),
      data.date        || "",
      data.time        || "",
      data.type        || "",
      data.address     || "",
      (data.unitsDeployed || []).join(", "),
      total,
      byUnit,
      data.narrative   || "",
      data.submittedBy || "Unknown"
    ]);

    return respond({ success: true });
  } catch(err) {
    return respond({ success: false, error: err.message });
  }
}

function doGet(e) {
  try {
    // Read calls
    const cs   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALLS_SHEET);
    const rows = [];
    if (cs) {
      const data = cs.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        const r = data[i];
        if (!r[3] && !r[4]) continue;
        rows.push({
          submittedAt:     String(r[0] || ""),
          date:            fmtDate(r[1]),
          time:            String(r[2] || ""),
          type:            String(r[3] || ""),
          address:         String(r[4] || ""),
          units:           String(r[5] || ""),
          personnel:       String(r[6] || "0"),
          personnelByUnit: String(r[7] || ""),
          narrative:       String(r[8] || ""),
          submittedBy:     String(r[9] || "Unknown"),
        });
      }
    }

    // Read config
    const cfg = readConfig();
    return respond({ rows, roster: cfg.roster, units: cfg.units });

  } catch(err) {
    return respond({ error: err.message });
  }
}

// ── CORS-friendly response ─────────────────────────────────────
// Apps Script doesn't support custom headers on doGet/doPost,
// but returning JSON with ContentService works cross-origin
// as long as the deployment is set to "Anyone" access.
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CONFIG ────────────────────────────────────────────────────
function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(CONFIG_SHEET);
  if (!s) {
    s = ss.insertSheet(CONFIG_SHEET);
    s.appendRow(["Type","Value"]);
    const h = s.getRange(1,1,1,2);
    h.setBackground("#1a1a2e"); h.setFontColor("#fff"); h.setFontWeight("bold");
    s.setColumnWidth(1,100); s.setColumnWidth(2,400);
  }
  return s;
}

function readConfig() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!s) return { roster:[], units:[] };
  const data = s.getDataRange().getValues();
  let roster=[], units=[];
  for (let i=1; i<data.length; i++) {
    const type = String(data[i][0]).trim();
    const val  = String(data[i][1]).trim();
    if (!val) continue;
    if (type==="roster") roster = val.split("|").map(x=>x.trim()).filter(Boolean);
    if (type==="units")  units  = val.split("|").map(x=>x.trim()).filter(Boolean);
  }
  return { roster, units };
}

function saveConfig(roster, units) {
  const s    = getConfigSheet();
  const data = s.getDataRange().getValues();
  let rRow=-1, uRow=-1;
  for (let i=1; i<data.length; i++) {
    if (data[i][0]==="roster") rRow=i+1;
    if (data[i][0]==="units")  uRow=i+1;
  }
  if (rRow>0) s.getRange(rRow,2).setValue(roster.join("|"));
  else s.appendRow(["roster", roster.join("|")]);
  if (uRow>0) s.getRange(uRow,2).setValue(units.join("|"));
  else s.appendRow(["units", units.join("|")]);
}

// ── CALLS SHEET ───────────────────────────────────────────────
function getCallsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(CALLS_SHEET);
  if (!s) {
    s = ss.insertSheet(CALLS_SHEET);
    s.appendRow(["Submitted At","Date","Time","Type","Address","Units","Personnel","Personnel by Unit","Narrative","Submitted By"]);
    const h = s.getRange(1,1,1,10);
    h.setBackground("#c0392b"); h.setFontColor("#fff"); h.setFontWeight("bold");
    s.setFrozenRows(1);
    [160,110,100,140,200,180,80,300,300,140].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }
  return s;
}

function fmtDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(val);
}
