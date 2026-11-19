// ==========================================
// RUANG CATATAN - BACKEND GAS V2
// ADMIN DASHBOARD + USER + CLOUD
// DATABASE LAMA TETAP DIPERTAHANKAN
// ==========================================

// ==========================================
// KONFIGURASI
// ==========================================

const SHEET_NAME = "Data";
const PROP_KEY = "CATATAN_SPREADSHEET_ID";
const SPREADSHEET_NAME = "Catatan Pribadi DB";

// GANTI DENGAN KUNCI ADMIN MILIKMU
// Gunakan kombinasi panjang dan sulit ditebak.
const ADMIN_KEY = "GANTI_DENGAN_KUNCI_ADMIN_RAHASIA_2026";

// STRUKTUR DATABASE:
// A = ID
// B = UpdatedAt
// C = JSON

// ==========================================
// DATABASE
// ==========================================

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

  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    ssId = ss.getId();

    props.setProperty(PROP_KEY, ssId);
  }

  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {

    sheet.getRange(1, 1, 1, 3).setValues([[
      "ID",
      "UpdatedAt",
      "JSON"
    ]]);

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 350);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 800);
  }

  return sheet;
}

// ==========================================
// URL SPREADSHEET
// ==========================================

function getSpreadsheetUrl() {

  const id = PropertiesService
    .getScriptProperties()
    .getProperty(PROP_KEY);

  return id
    ? "https://docs.google.com/spreadsheets/d/" + id
    : "";
}

// ==========================================
// RESPONSE JSON / JSONP
// ==========================================

function jsonResponse(data, callback) {

  const output = ContentService.createTextOutput();

  if (callback) {

    // Batasi karakter callback untuk mencegah
    // penyalahgunaan nama fungsi JSONP.
    if (!/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback)) {
      throw new Error("Callback tidak valid.");
    }

    output.setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );

    output.setContent(
      callback + "(" + JSON.stringify(data) + ");"
    );

  } else {

    output.setMimeType(
      ContentService.MimeType.JSON
    );

    output.setContent(
      JSON.stringify(data)
    );
  }

  return output;
}

// ==========================================
// CARI BARIS BERDASARKAN ID
// ==========================================

function findRowById(sheet, id) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return -1;

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < ids.length; i++) {

    if (
      String(ids[i][0]).trim() === String(id).trim()
    ) {
      return i + 2;
    }
  }

  return -1;
}

// ==========================================
// NORMALISASI DATA AKUN
// ==========================================

function normalizeAccount(id, data) {

  return {

    version: 1,

    id: String(id),

    notes: Array.isArray(data.notes)
      ? data.notes.map(n => ({

          id: String(n.id || ""),

          title: String(n.title || ""),

          body: String(n.body || ""),

          visibility:
            n.visibility === "public"
              ? "public"
              : "private",

          updatedAt:
            Number(n.updatedAt) || Date.now()

        }))
      : [],

    updatedAt:
      Number(data.updatedAt) || Date.now()

  };
}

// ==========================================
// BACA DATA JSON
// ==========================================

function readAccount(jsonString, id) {

  const data = JSON.parse(jsonString);

  return normalizeAccount(id, data);
}

// ==========================================
// ADMIN: AMBIL SEMUA DATABASE
// ==========================================

function getAllAdminData(sheet, key) {

  // Validasi kunci admin di server.
  if (!ADMIN_KEY || ADMIN_KEY.indexOf("GANTI_") === 0) {
    throw new Error("ADMIN KEY belum dikonfigurasi.");
  }

  if (String(key || "") !== ADMIN_KEY) {
    throw new Error("ADMIN KEY salah atau tidak valid.");
  }

  const lastRow = sheet.getLastRow();

  const users = [];

  let totalNotes = 0;
  let totalPublic = 0;
  let totalPrivate = 0;

  if (lastRow >= 2) {

    // Baca semua kolom A sampai C.
    const rows = sheet
      .getRange(2, 1, lastRow - 1, 3)
      .getValues();

    rows.forEach(row => {

      const id = String(row[0] || "").trim();

      const updatedAt = row[1];

      const jsonString = row[2];

      if (!id || !jsonString) return;

      try {

        const data = readAccount(jsonString, id);

        const notes = data.notes;

        const noteCount = notes.length;

        const publicCount = notes.filter(
          n => n.visibility === "public"
        ).length;

        const privateCount = notes.filter(
          n => n.visibility !== "public"
        ).length;

        totalNotes += noteCount;
        totalPublic += publicCount;
        totalPrivate += privateCount;

        // Seluruh data akun dikirim ke dashboard.
        // Termasuk catatan privat.
        users.push({

          id: id,

          updatedAt:
            Number(updatedAt) ||
            data.updatedAt ||
            0,

          noteCount: noteCount,

          publicCount: publicCount,

          privateCount: privateCount,

          data: data

        });

      } catch (err) {

        // Baris rusak tidak menghentikan
        // pembacaan akun lainnya.
        users.push({

          id: id,

          updatedAt: updatedAt || 0,

          noteCount: 0,

          publicCount: 0,

          privateCount: 0,

          data: {
            id: id,
            notes: [],
            error: "JSON tidak dapat dibaca: " + err.message
          }

        });

      }

    });

  }

  users.sort((a, b) => {

    return Number(b.updatedAt) - Number(a.updatedAt);

  });

  return {

    ok: true,

    users: users,

    totalUsers: users.length,

    totalNotes: totalNotes,

    totalPublic: totalPublic,

    totalPrivate: totalPrivate,

    serverTime: Date.now()

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

    // ======================================
    // 1. AMBIL SATU AKUN
    // ======================================

    if (action === "get") {

      const id = String(p.id || "").trim();

      if (!/^\d{25,45}$/.test(id)) {

        return jsonResponse({
          ok: false,
          error: "ID tidak valid."
        }, callback);

      }

      const row = findRowById(sheet, id);

      if (row === -1) {

        return jsonResponse({

          ok: true,

          found: false,

          data: null

        }, callback);

      }

      const jsonString = sheet
        .getRange(row, 3)
        .getValue();

      const data = readAccount(jsonString, id);

      return jsonResponse({

        ok: true,

        found: true,

        data: data

      }, callback);

    }

    // ======================================
    // 2. AMBIL CATATAN PUBLIK
    // ======================================

    if (action === "public") {

      const lastRow = sheet.getLastRow();

      const notes = [];

      if (lastRow >= 2) {

        const rows = sheet
          .getRange(2, 1, lastRow - 1, 3)
          .getValues();

        rows.forEach(row => {

          const userId = String(row[0] || "");

          const jsonString = row[2];

          if (!jsonString) return;

          try {

            const data = JSON.parse(jsonString);

            if (Array.isArray(data.notes)) {

              data.notes.forEach(n => {

                if (n.visibility === "public") {

                  notes.push({

                    id: String(n.id || ""),

                    title: String(n.title || ""),

                    body: String(n.body || ""),

                    visibility: "public",

                    updatedAt:
                      Number(n.updatedAt) || 0

                  });

                }

              });

            }

          } catch (err) {}

        });

      }

      notes.sort((a, b) => {

        return b.updatedAt - a.updatedAt;

      });

      return jsonResponse({

        ok: true,

        notes: notes.slice(0, 100)

      }, callback);

    }

    // ======================================
    // 3. ADMIN DASHBOARD
    // ======================================

    if (action === "admin") {

      const key = String(p.key || "");

      const result = getAllAdminData(sheet, key);

      return jsonResponse(result, callback);

    }

    // ======================================
    // 4. INFO DATABASE
    // ======================================

    if (action === "info") {

      return jsonResponse({

        ok: true,

        sheetUrl: getSpreadsheetUrl(),

        database: SHEET_NAME

      }, callback);

    }

    // ======================================
    // ACTION TIDAK DIKENAL
    // ======================================

    return jsonResponse({

      ok: false,

      error: "Action tidak dikenal."

    }, callback);

  } catch (err) {

    return jsonResponse({

      ok: false,

      error: err.message || String(err)

    }, (e.parameter && e.parameter.callback) || "");

  }

}

// ==========================================
// POST - SIMPAN DATA PENGGUNA
// ==========================================

function doPost(e) {

  try {

    let payload = null;

    if (e.parameter && e.parameter.payload) {

      payload = JSON.parse(e.parameter.payload);

    } else if (
      e.postData &&
      e.postData.contents
    ) {

      payload = JSON.parse(e.postData.contents);

    } else {

      throw new Error("Payload kosong.");

    }

    if (payload.action !== "save") {

      throw new Error("Action POST tidak dikenal.");

    }

    const id = String(payload.id || "").trim();

    if (!/^\d{25,45}$/.test(id)) {

      throw new Error("ID tidak valid.");

    }

    if (
      !payload.data ||
      !Array.isArray(payload.data.notes)
    ) {

      throw new Error("Data akun tidak valid.");

    }

    const toSave = normalizeAccount(
      id,
      payload.data
    );

    const jsonString = JSON.stringify(toSave);

    const lock = LockService.getScriptLock();

    lock.waitLock(15000);

    try {

      const sheet = getSheet();

      const row = findRowById(sheet, id);

      if (row !== -1) {

        // Update kolom B dan C.
        sheet.getRange(row, 2, 1, 2).setValues([[
          toSave.updatedAt,
          jsonString
        ]]);

      } else {

        // Tambahkan akun baru.
        sheet.appendRow([
          id,
          toSave.updatedAt,
          jsonString
        ]);

      }

      SpreadsheetApp.flush();

    } finally {

      lock.releaseLock();

    }

    return ContentService
      .createTextOutput(JSON.stringify({

        ok: true,

        id: id,

        updatedAt: toSave.updatedAt

      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({

        ok: false,

        error: err.message || String(err)

      }))
      .setMimeType(ContentService.MimeType.JSON);

  }

}

// ==========================================
// CEK DATABASE
// JALANKAN MANUAL JIKA DIPERLUKAN
// ==========================================

function cekSheetUrl() {

  getSheet();

  Logger.log(getSpreadsheetUrl());

}

// ==========================================
// TES ADMIN
// JALANKAN MANUAL DI EDITOR GAS
// ==========================================

function testAdminDatabase() {

  const sheet = getSheet();

  const result = getAllAdminData(
    sheet,
    ADMIN_KEY
  );

  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;

}