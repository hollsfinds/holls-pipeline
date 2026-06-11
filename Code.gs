/**
 * Holls Finds Pipeline — Sheet backend (Google Apps Script Web App)
 * Read/write JSON bridge so the web board can use the Sheet without Cowork.
 *
 * Deploy: Extensions > Apps Script (from the Pipeline Sheet) > paste this >
 *   Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
 *   (Keep the SAME web app URL so the board doesn't change.)
 *
 * Security: obscure URL + shared TOKEN below. The web app uses this same token.
 */

const SHEET_ID = "1vKbLGryulWJ-dKV6YVI6jZ9kg0Qs9U1_OxK6hiDpfvI";
const TOKEN    = "holls-pipeline-2026";                 // must match API_TOKEN in the web app
const TABS     = ["Inbox", "Pipeline", "Stats", "Social", "On Deck", "Meta", "NonPromo", "Calendar", "Roadmap"];  // On Deck = product intake staging; Meta = category/asin cache; Roadmap = revenue roadmap (board tab)

// On Deck column order (1-based). Keep in sync with the Sheet header row.
// A Date Added | B ASIN | C Cleaned Title | D Raw Title | E Image URL |
// F Product URL | G Categories | H New Categories | I Status |
// J FreshStore Product ID | K Notes / Error
const ONDECK_TAB = "On Deck";

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) return json({ error: "unauthorized" });

  // Category picker / dedupe cache for the bookmarklet.
  // Returns the Meta tab as a key->value map (e.g. freshstore_categories, freshstore_asins).
  // The Mac brain keeps these fresh; missing Meta tab just returns {}.
  if (e.parameter.action === "meta") {
    return json({ ok: true, meta: readMeta() });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  // Optional ?tabs=Stats,Pipeline filter — returns ONLY those tabs (must be in TABS).
  // Lightweight polls (heartbeat/pause pill, sweep gate) use this; no param = all tabs (back-compat).
  let want = TABS;
  if (e.parameter.tabs) {
    const req = String(e.parameter.tabs).split(",").map(s => s.trim()).filter(Boolean);
    const allowed = req.filter(t => TABS.indexOf(t) !== -1);
    if (allowed.length) want = allowed;
  }
  const data = {};
  want.forEach(t => {
    const sh = ss.getSheetByName(t);
    data[t] = sh ? sh.getDataRange().getValues() : [];
  });
  return json({ ok: true, data: data });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json({ error: "bad json" }); }
  if (b.token !== TOKEN) return json({ error: "unauthorized" });

  // LockService: the board + the Mac brain both write; serialize so concurrent
  // writes can't race. 10s wait is generous (writes take <1s).
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json({ error: "busy — write lock timeout" }); }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  try {
    // --- On Deck intake from the bookmarklet (deduped by ASIN) ---
    if (b.action === "addOnDeck") {
      return json(addOnDeck(ss, b));
    }

    // --- generic Sheet actions (used by the board) ---
    let sh = ss.getSheetByName(b.sheet);
    if (!sh) {
      if (b.createIfMissing) sh = ss.insertSheet(b.sheet);   // e.g. the Calendar tab self-creates on first write
      else return json({ error: "no sheet: " + b.sheet });
    }

    if (b.action === "updateCell") {            // {row, col, value} (1-based)
      sh.getRange(b.row, b.col).setValue(b.value);
      return json({ ok: true });
    }
    if (b.action === "updateRange") {           // {range:"A2:F2", values:[[...]]}
      sh.getRange(b.range).setValues(b.values);
      return json({ ok: true });
    }
    if (b.action === "appendRow") {             // {values:[...]}
      sh.appendRow(b.values);
      return json({ ok: true, row: sh.getLastRow() });
    }
    if (b.action === "batch") {                 // {ops:[{range:"C5",values:[[...]]},...]} — several writes, ONE round trip (same sheet)
      if (!Array.isArray(b.ops) || !b.ops.length) return json({ error: "batch: no ops" });
      b.ops.forEach(op => { sh.getRange(op.range).setValues(op.values); });
      return json({ ok: true, applied: b.ops.length });
    }
    return json({ error: "unknown action: " + b.action });
  } catch (err) {
    return json({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Add a bookmarked Amazon product to the On Deck tab, deduped by ASIN.
 * Payload: { asin, cleanedTitle, rawTitle, imageUrl, productUrl,
 *            categories:[...], newCategories:[...] }
 * Returns: { ok:true, status:"added", row:n }
 *       or { ok:true, status:"duplicate", where:"on-deck"|"freshstore" }
 *       or { error:"..." }
 */
function addOnDeck(ss, b) {
  const asin = String(b.asin || "").trim().toUpperCase();
  if (!asin) return { error: "missing asin" };

  const sh = ss.getSheetByName(ONDECK_TAB);
  if (!sh) return { error: "no On Deck tab" };

  // 1. dedupe against existing On Deck rows (column B = ASIN)
  const last = sh.getLastRow();
  if (last >= 2) {
    const existing = sh.getRange(2, 2, last - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim().toUpperCase() === asin) {
        return { ok: true, status: "duplicate", where: "on-deck" };
      }
    }
  }

  // 2. dedupe against live FreshStore products (ASIN cache maintained by the Mac brain)
  const meta = readMeta();
  if (meta.freshstore_asins) {
    let live = [];
    try { live = JSON.parse(meta.freshstore_asins); } catch (err) { live = []; }
    if (live.map(String).map(s => s.toUpperCase()).indexOf(asin) !== -1) {
      return { ok: true, status: "duplicate", where: "freshstore" };
    }
  }

  // 3. append a new On Deck row (Status = New)
  const cats    = (b.categories    || []).join(", ");
  const newCats = (b.newCategories || []).join(", ");
  const row = [
    new Date(),                 // A Date Added
    asin,                       // B ASIN
    b.cleanedTitle || "",       // C Cleaned Title (editable on the card)
    b.rawTitle     || "",       // D Raw Title (fallback)
    b.imageUrl     || "",       // E Image URL
    b.productUrl   || ("https://www.amazon.com/dp/" + asin), // F Product URL
    cats,                       // G Categories
    newCats,                    // H New Categories
    "New",                      // I Status
    "",                         // J FreshStore Product ID
    ""                          // K Notes / Error
  ];
  sh.appendRow(row);
  return { ok: true, status: "added", row: sh.getLastRow() };
}

/** Read the optional "Meta" tab as a key->value map (col A = key, col B = value). */
function readMeta() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Meta");
  const out = {};
  if (!sh) return out;
  const vals = sh.getDataRange().getValues();
  vals.forEach(r => { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
