
/**
 * CATATAN PRIBADI CLOUD V1.1
 * STANDALONE GOOGLE APPS SCRIPT
 * Database dibuat otomatis.
 */

const CONFIG = {
  SHEET_NAME: "Users",
  DB_PROPERTY: "DB_ID",
  ADMIN_PROPERTY: "ADMIN_KEY",
  MAX_SIZE: 450000
};

// =================================
// 1. SETUP OTOMATIS
// =================================

function setup() {
  const props = PropertiesService.getScriptProperties();

  // Buat database Google Sheets otomatis
  let dbId = props.getProperty(CONFIG.DB_PROPERTY);

  if (!dbId) {
    const ss = SpreadsheetApp.create(
      "DATABASE CATATAN PRIBADI CLOUD"
    );

    const sh = ss.getSheets()[0];
    sh.setName(CONFIG.SHEET_NAME);

    sh.appendRow([
      "ID",
      "DATA_JSON",
      "UPDATED_AT",
      "NOTE_COUNT"
    ]);

    dbId = ss.getId();
    props.setProperty(CONFIG.DB_PROPERTY, dbId);
  }

  // Buat kunci administrator
  let adminKey = props.getProperty(CONFIG.ADMIN_PROPERTY);

  if (!adminKey) {
    adminKey = Utilities.getUuid() +
               Utilities.getUuid();

    props.setProperty(CONFIG.ADMIN_PROPERTY, adminKey);
  }

  Logger.log("SETUP BERHASIL");
  Logger.log("DATABASE: https://docs.google.com/spreadsheets/d/" + dbId);
  Logger.log("ADMIN KEY: " + adminKey);
}

// =================================
// 2. DATABASE
// =================================

function getSheet_() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty(CONFIG.DB_PROPERTY);

  if (!id) throw new Error("Jalankan setup terlebih dahulu.");

  return SpreadsheetApp
    .openById(id)
    .getSheetByName(CONFIG.SHEET_NAME);
}

function validId_(id) {
  return /^\d{25,45}$/.test(String(id || ""));
}

function findUser_(sh, id) {
  const last = sh.getLastRow();

  if (last < 2) return 0;

  const found = sh.getRange(2, 1, last - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();

  return found ? found.getRow() : 0;
}

// =================================
// 3. JSONP RESPONSE
// =================================

function output_(data, callback) {
  const cb = /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback || "")
    ? callback
    : "syncCallback";

  return ContentService
    .createTextOutput(
      cb + "(" + JSON.stringify(data) + ");"
    )
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// =================================
// 4. GET DATA
// =================================

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action;
    const sh = getSheet_();

    // LOGIN / AMBIL CATATAN
    if (action === "get") {
      const id = String(p.id || "");

      if (!validId_(id)) {
        throw new Error("ID tidak valid.");
      }

      const row = findUser_(sh, id);

      if (!row) {
        return output_({
          ok: true,
          found: false
        }, p.callback);
      }

      const data = JSON.parse(
        sh.getRange(row, 2).getValue()
      );

      return output_({
        ok: true,
        found: true,
        data: data
      }, p.callback);
    }

    // CATATAN PUBLIK
    if (action === "public") {
      const values = sh.getDataRange().getValues();
      const notes = [];

      for (let i = 1; i < values.length; i++) {
        try {
          const data = JSON.parse(values[i][1]);

          (data.notes || []).forEach(n => {
            if (n.visibility === "public") {
              notes.push({
                title: String(n.title || ""),
                body: String(n.body || ""),
                updatedAt: Number(n.updatedAt) || 0
              });
            }
          });
        } catch (err) {}
      }

      notes.sort((a,b) => b.updatedAt - a.updatedAt);

      return output_({
        ok: true,
        notes: notes.slice(0, 300)
      }, p.callback);
    }

    // DASHBOARD ADMIN
    if (action === "admin") {
      const key = PropertiesService
        .getScriptProperties()
        .getProperty(CONFIG.ADMIN_PROPERTY);

      if (!key || p.key !== key) {
        throw new Error("Kunci admin salah.");
      }

      const values = sh.getDataRange().getValues();
      const users = [];
      let totalNotes = 0;

      for (let i = 1; i < values.length; i++) {
        const count = Number(values[i][3]) || 0;
        totalNotes += count;

        users.push({
          id: String(values[i][0]),
          noteCount: count,
          updatedAt: values[i][2]
            ? new Date(values[i][2]).getTime()
            : 0
        });
      }

      users.sort((a,b) => b.updatedAt - a.updatedAt);

      return output_({
        ok: true,
        users: users,
        totalNotes: totalNotes
      }, p.callback);
    }

    throw new Error("Perintah tidak dikenal.");

  } catch (err) {
    return output_({
      ok: false,
      error: String(err.message || err)
    }, e && e.parameter ? e.parameter.callback : "");
  }
}

// =================================
// 5. SAVE TO CLOUD
// =================================

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const body = JSON.parse(e.postData.contents || "{}");

    if (body.action !== "save") {
      throw new Error("Perintah tidak valid.");
    }

    const id = String(body.id || "");

    if (!validId_(id)) {
      throw new Error("ID tidak valid.");
    }

    const data = body.data;

    if (
      !data ||
      String(data.id) !== id ||
      !Array.isArray(data.notes)
    ) {
      throw new Error("Format catatan tidak valid.");
    }

    const json = JSON.stringify(data);

    if (json.length > CONFIG.MAX_SIZE) {
      throw new Error("Ukuran data terlalu besar.");
    }

    if (data.notes.length > 5000) {
      throw new Error("Maksimal 5000 catatan.");
    }

    const sh = getSheet_();
    const row = findUser_(sh, id);

    const values = [[
      id,
      json,
      new Date(),
      data.notes.length
    ]];

    if (row) {
      sh.getRange(row, 1, 1, 4).setValues(values);
    } else {
      sh.appendRow(values[0]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        message: "Data tersimpan."
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(err.message || err)
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
