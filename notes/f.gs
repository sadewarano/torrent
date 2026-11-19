// ==========================================
// CATATAN PRIBADI - BACKEND GAS VERSI SHEET
// Auto-create Spreadsheet + Sheet
// ==========================================

const SHEET_NAME = "Data";
const PROP_KEY = "CATATAN_SPREADSHEET_ID";
const SPREADSHEET_NAME = "Catatan Pribadi DB";

// Header: A=ID | B=UpdatedAt | C=JSON (isi akun lengkap)
function getSheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(PROP_KEY);
  let ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = null;
    }
  }

  // Kalau belum ada / kehapus, bikin baru otomatis
  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    ssId = ss.getId();
    props.setProperty(PROP_KEY, ssId);
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Buat header kalau sheet masih kosong
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([["ID", "UpdatedAt", "JSON"]]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 350);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 800);
  }

  return sheet;
}

function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_KEY);
  return id ? "https://docs.google.com/spreadsheets/d/" + id : "";
}

function jsonResponse(data, callback) {
  const output = ContentService.createTextOutput();
  if (callback) {
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    output.setContent(callback + "(" + JSON.stringify(data) + ");");
  } else {
    output.setMimeType(ContentService.MimeType.JSON);
    output.setContent(JSON.stringify(data));
  }
  return output;
}

// Cari baris berdasarkan ID, return nomor baris atau -1
function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  // Ambil kolom A saja biar cepat
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function normalizeAccount(id, data) {
  return {
    version: 1,
    id: String(id),
    notes: Array.isArray(data.notes) ? data.notes.map(n => ({
      id: String(n.id || ""),
      title: String(n.title || ""),
      body: String(n.body || ""),
      visibility: n.visibility === "public" ? "public" : "private",
      updatedAt: Number(n.updatedAt) || Date.now()
    })) : [],
    updatedAt: Number(data.updatedAt) || Date.now()
  };
}

// ==========================================
// GET
// ==========================================
function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || "";
    const callback = p.callback || "";
    const sheet = getSheet();

    // 1. GET 1 AKUN
    if (action === "get") {
      const id = String(p.id || "").trim();
      if (!id) return jsonResponse({ ok: false, error: "ID kosong" }, callback);

      const row = findRowById(sheet, id);
      if (row === -1) {
        return jsonResponse({ ok: true, found: false, data: null }, callback);
      }

      const jsonStr = sheet.getRange(row, 3).getValue();
      const data = JSON.parse(jsonStr);
      return jsonResponse({ ok: true, found: true, data: data }, callback);
    }

    // 2. GET PUBLIK
    if (action === "public") {
      const lastRow = sheet.getLastRow();
      const notes = [];
      if (lastRow >= 2) {
        const rows = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        rows.forEach(r => {
          try {
            const data = JSON.parse(r[0]);
            if (Array.isArray(data.notes)) {
              data.notes.forEach(n => {
                if (n.visibility === "public") {
                  notes.push({
                    id: String(n.id || ""),
                    title: String(n.title || ""),
                    body: String(n.body || ""),
                    visibility: "public",
                    updatedAt: Number(n.updatedAt) || 0
                  });
                }
              });
            }
          } catch (err) {}
        });
      }
      notes.sort((a, b) => b.updatedAt - a.updatedAt);
      return jsonResponse({ ok: true, notes: notes.slice(0, 100) }, callback);
    }

    // 3. INFO - buat cek URL sheet-nya
    if (action === "info") {
      return jsonResponse({ ok: true, sheetUrl: getSpreadsheetUrl() }, callback);
    }

    return jsonResponse({ ok: false, error: "Action tidak dikenal" }, callback);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() }, (e.parameter && e.parameter.callback) || "");
  }
}

// ==========================================
// POST - SAVE
// ==========================================
function doPost(e) {
  try {
    let payload = null;
    if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      throw new Error("Payload kosong");
    }

    if (payload.action !== "save") throw new Error("Action POST tidak dikenal");

    const id = String(payload.id || "").trim();
    if (!/^\d{25,45}$/.test(id)) throw new Error("ID tidak valid");
    if (!payload.data || !Array.isArray(payload.data.notes)) throw new Error("Data tidak valid");

    const toSave = normalizeAccount(id, payload.data);
    const sheet = getSheet();
    const jsonStr = JSON.stringify(toSave);

    const row = findRowById(sheet, id);
    if (row !== -1) {
      // Update baris yang ada
      sheet.getRange(row, 2, 1, 2).setValues([[toSave.updatedAt, jsonStr]]);
    } else {
      // Tambah baris baru
      sheet.appendRow([id, toSave.updatedAt, jsonStr]);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// FUNGSI BANTUAN - Jalankan sekali di editor
// untuk lihat URL Sheet-nya di Log
// ==========================================
function cekSheetUrl() {
  getSheet();
  Logger.log(getSpreadsheetUrl());
}