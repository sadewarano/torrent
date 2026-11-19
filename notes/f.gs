/**
 * CATATAN PRIBADI CLOUD
 * Google Apps Script backend
 *
 * Buat spreadsheet, buka Extensions > Apps Script,
 * tempel kode ini, simpan, lalu jalankan setup().
 */

const CONFIG = {
  SHEET_NAME: "Users",
  MAX_DATA_CHARS: 450000
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Buat script dari spreadsheet: Extensions > Apps Script.");

  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  sh.clear();
  sh.appendRow(["ID", "DATA_JSON", "UPDATED_AT", "NOTE_COUNT"]);

  PropertiesService.getScriptProperties().setProperty("DB_ID", ss.getId());
  const adminKey = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty("ADMIN_KEY", adminKey);
  Logger.log("KUNCI ADMIN (simpan rahasia): " + adminKey);
  Logger.log("Setup selesai. Spreadsheet ID: " + ss.getId());
}

function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("DB_ID");
  if (!id) throw new Error("Jalankan setup() terlebih dahulu.");
  const sh = SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error("Sheet database tidak ditemukan.");
  return sh;
}

function validId_(id) {
  return /^\d{25,45}$/.test(String(id || ""));
}

function output_(data, callback) {
  const cb = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback || "") ? callback : "syncCallback";
  return ContentService.createTextOutput(cb + "(" + JSON.stringify(data) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = String(p.action || "");
    const sh = sheet_();

    if (action === "get") {
      const id = String(p.id || "");
      if (!validId_(id)) return output_({ok:false,error:"ID tidak valid"}, p.callback);
      const row = findUser_(sh, id);
      if (!row) return output_({ok:true,found:false}, p.callback);

      const raw = sh.getRange(row, 2).getValue();
      let data;
      try { data = JSON.parse(raw); }
      catch (_) { return output_({ok:false,error:"Data akun rusak"}, p.callback); }

      return output_({ok:true,found:true,data:data,updatedAt:sh.getRange(row,3).getValue()}, p.callback);
    }

    if (action === "admin") {
      const expected = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
      if (!expected || String(p.key || "") !== expected) {
        return output_({ok:false,error:"Kunci admin salah"}, p.callback);
      }
      const values = sh.getDataRange().getValues();
      const users = [];
      let totalNotes = 0;
      for (let i = 1; i < values.length; i++) {
        let count = Number(values[i][3]) || 0;
        totalNotes += count;
        users.push({
          id: String(values[i][0]),
          noteCount: count,
          updatedAt: values[i][2] ? new Date(values[i][2]).getTime() : 0
        });
      }
      users.sort((a,b) => b.updatedAt - a.updatedAt);
      return output_({ok:true,users:users,totalNotes:totalNotes}, p.callback);
    }

    if (action === "public") {
      const values = sh.getDataRange().getValues();
      const notes = [];
      for (let i = 1; i < values.length; i++) {
        try {
          const account = JSON.parse(values[i][1]);
          (account.notes || []).forEach(n => {
            if (n && n.visibility === "public") {
              notes.push({
                title:String(n.title || ""),
                body:String(n.body || ""),
                updatedAt:Number(n.updatedAt) || 0
              });
            }
          });
        } catch (_) {}
      }
      notes.sort((a,b) => b.updatedAt - a.updatedAt);
      return output_({ok:true,notes:notes.slice(0,300)}, p.callback);
    }

    return output_({ok:false,error:"Action tidak dikenal"}, p.callback);
  } catch (err) {
    return output_({ok:false,error:String(err.message || err)}, e?.parameter?.callback);
  }
}

function findUser_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const found = sh.getRange(2, 1, last - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents || "{}");

    if (body.action !== "save") throw new Error("Action tidak valid.");
    const id = String(body.id || "");
    if (!validId_(id)) throw new Error("ID tidak valid.");

    const data = body.data;
    if (!data || String(data.id) !== id || !Array.isArray(data.notes)) {
      throw new Error("Struktur data tidak valid.");
    }

    const json = JSON.stringify(data);
    if (json.length > CONFIG.MAX_DATA_CHARS) {
      throw new Error("Ukuran catatan terlalu besar.");
    }
    if (data.notes.length > 5000) {
      throw new Error("Maksimal 5000 catatan per akun.");
    }

    const sh = sheet_();
    const row = findUser_(sh, id);
    const now = new Date();
    const values = [[id, json, now, data.notes.length]];

    if (row) {
      sh.getRange(row, 1, 1, 4).setValues(values);
    } else {
      sh.appendRow(values[0]);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok:false,error:String(err.message || err)
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
