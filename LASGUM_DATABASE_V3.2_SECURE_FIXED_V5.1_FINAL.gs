/************************************************************
 * LASGUM DATABASE API
 * VERSION 3.2 FINAL
 *
 * Offline First
 * Google Apps Script + Google Spreadsheet
 *
 * Schema Version : 5
 ************************************************************/

const LASGUM_SCHEMA_VERSION = 5;
const LASGUM_RELEASE = "5.1-FINAL";


/* =========================================================
   NAMA DAN STRUKTUR DATABASE
   ========================================================= */

const LASGUM_SHEETS = {

  SEKOLAH: [
    "id","nama","npsn","alamat","desaKecamatan",
    "kabupatenProvinsi","kepalaSekolah","logoUrl",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  PENGGUNA: [
    "id","username","nama","role","passwordHash","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  GURU: [
    "id","penggunaId","nik","nip","nuptk","nama",
    "jenisKelamin","nomorHp","email","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  SISWA: [
    "id","nis","nisn","nama","jenisKelamin",
    "tempatLahir","tanggalLahir","kelas","rombel","fase",
    "status","createdAt","updatedAt","deleted",
    "deviceId","schemaVersion","ownerId","asalSekolah"
  ],

  MATA_PELAJARAN: [
    "id","kode","nama","kelas","fase","cakupan","scopeVersion","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  MATERI_CP: [
    "id","mataPelajaranId","kelas","fase","kodeCp",
    "cp","material","deskripsi","semester","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  SOAL: [
    "id","mataPelajaranId","materiId","kelas","fase",
    "tipe","pertanyaan","opsi","kunci","kataKunci",
    "pembahasan","bobot","status","sumber","createdBy",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  PAKET_UJIAN: [
    "id","kode","nama","mataPelajaranId","kelas","fase",
    "durasiMenit","jumlahSoal","acakSoal","acakOpsi",
    "dipublikasikan","status","createdBy",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  PAKET_SOAL: [
    "id","paketId","soalId","nomor","bobot",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  JADWAL_UJIAN: [
    "id","paketId","mulai","selesai","aktif","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  UJIAN: [
    "id","paketId","siswaId","mulai","selesai",
    "status","attempt","deviceId","createdAt",
    "updatedAt","deleted","schemaVersion"
  ],

  JAWABAN: [
    "id","ujianId","soalId","jawaban","benar","skor",
    "manual","createdAt","updatedAt","deleted",
    "deviceId","schemaVersion"
  ],

  NILAI: [
    "id","ujianId","siswaId","paketId",
    "skorPg","skorIsian","skorUraian",
    "nilaiAkhir","status","catatan",
    "createdAt","updatedAt","deleted",
    "deviceId","schemaVersion"
  ],

  PENGATURAN: [
    "id","kunci","nilai","deskripsi",
    "updatedAt","schemaVersion"
  ],

  SYNC_LOG: [
    "id","deviceId","action","tableName","recordId",
    "clientUpdatedAt","serverUpdatedAt","status",
    "message","createdAt","schemaVersion"
  ]

};


/* =========================================================
   UTILITAS
   ========================================================= */

function lasgumNow_() {
  return new Date().toISOString();
}


function lasgumId_() {
  return Utilities.getUuid();
}


function lasgumJson_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


function lasgumParse_(e) {

  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error("JSON request tidak valid.");
  }
}


function lasgumGetSheet_(tableName) {

  if (!LASGUM_SHEETS[tableName]) {
    throw new Error(
      "Tabel tidak dikenal: " + tableName
    );
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(tableName);

  if (!sheet) {
    throw new Error(
      "Sheet tidak ditemukan: " + tableName
    );
  }

  return sheet;
}


/* =========================================================
   SETUP DATABASE
   ========================================================= */

function setupDatabase() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(LASGUM_SHEETS).forEach(function(name) {
    ensureSheetSchema_(ss, name, LASGUM_SHEETS[name]);
  });

  setupDefaultSettings_();
  SpreadsheetApp.flush();

  return {
    ok: true,
    schemaVersion: LASGUM_SCHEMA_VERSION,
    database: ss.getName(),
    sheets: Object.keys(LASGUM_SHEETS),
    serverTime: lasgumNow_()
  };
}

/**
 * Memastikan header dan posisi kolom sesuai schema tanpa merusak data lama.
 * Jika kolom baru ditambahkan, data dipindahkan berdasarkan nama header lama,
 * bukan berdasarkan nomor kolom. Ini penting untuk migrasi V4 -> V5.
 */
function ensureSheetSchema_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const oldHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(function(v){ return String(v || '').trim(); });
    const oldIndex = {};
    oldHeaders.forEach(function(h, i){ if (h) oldIndex[h] = i; });

    const oldRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    const newRows = oldRows.map(function(oldRow) {
      return headers.map(function(h) {
        return Object.prototype.hasOwnProperty.call(oldIndex, h) ? oldRow[oldIndex[h]] : '';
      });
    });

    // Tulis header + data pada layout baru.
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (newRows.length) {
      sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
    }
  }

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
  sheet.setFrozenRows(1);
}


function setupDefaultSettings_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("PENGATURAN");

  if (!sheet) return;

  const existing = {};

  if (sheet.getLastRow() > 1) {

    const values =
      sheet
        .getRange(
          2,
          2,
          sheet.getLastRow() - 1,
          1
        )
        .getValues();

    values.forEach(function(row) {

      if (row[0]) {
        existing[String(row[0])] = true;
      }

    });

  }


  const defaults = [

    [
      "schemaVersion",
      LASGUM_SCHEMA_VERSION,
      "Versi struktur database"
    ],

    [
      "databaseName",
      "DATABASE LASGUM",
      "Nama database LASGUM"
    ],

    [
      "syncMode",
      "incremental",
      "Mode sinkronisasi bertahap"
    ],

    [
      "offlineFirst",
      "true",
      "Aplikasi mendukung penggunaan offline"
    ],

    [
      "backupMode",
      "incremental",
      "Backup berdasarkan perubahan data"
    ]

  ];


  const rows = defaults
    .filter(function(item) {
      return !existing[item[0]];
    })
    .map(function(item) {

      return [
        lasgumId_(),
        item[0],
        item[1],
        item[2],
        lasgumNow_(),
        LASGUM_SCHEMA_VERSION
      ];

    });


  if (rows.length) {

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        6
      )
      .setValues(rows);

  }

}


/* =========================================================
   HEALTH
   ========================================================= */

function healthCheck() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  return {

    ok: true,

    database: ss.getName(),

    schemaVersion:
      LASGUM_SCHEMA_VERSION,

    serverTime:
      lasgumNow_(),

    message:
      "LASGUM Database API aktif."

  };

}


/* =========================================================
   GET DATA
   ========================================================= */

function apiGet_(tableName, params) {

  const sheet =
    lasgumGetSheet_(tableName);

  const headers =
    LASGUM_SHEETS[tableName];

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return {
      ok: true,
      table: tableName,
      count: 0,
      data: []
    };

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const data = values.map(function(row) {

    const obj = {};

    headers.forEach(function(header,index) {

      obj[header] =
        normalizeValue_(row[index], header, tableName);

    });

    return obj;

  });


  /*
   * Secara default data deleted=true tidak ditampilkan.
   */

  let filtered =
    data.filter(function(row) {

      return String(row.deleted)
        .toLowerCase() !== "true";

    });


  /*
   * Filter berdasarkan id.
   */

  if (params && params.id) {

    filtered =
      filtered.filter(function(row) {

        return String(row.id) ===
          String(params.id);

      });

  }


  return {

    ok: true,

    table: tableName,

    count: filtered.length,

    data: filtered,

    serverTime: lasgumNow_(),

    schemaVersion:
      LASGUM_SCHEMA_VERSION

  };

}


/* =========================================================
   CREATE
   ========================================================= */

function apiCreate_(tableName, record, deviceId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = lasgumGetSheet_(tableName);
    const headers = LASGUM_SHEETS[tableName];
    record = normalizeRecord_(tableName, record || {});

    if (!record.id) {
      record.id = lasgumId_();
    }

    /*
     * CREATE idempotent:
     * Jika perangkat mengirim ulang record yang sama karena
     * koneksi terputus/retry, jangan membuat baris duplikat.
     */
    const existingRow = findRowById_(sheet, headers, record.id);
    if (existingRow !== -1) {
      const existing = readRowObject_(sheet, headers, existingRow);
      return {
        ok: true,
        action: "CREATE",
        table: tableName,
        duplicate: true,
        data: existing,
        serverTime: lasgumNow_()
      };
    }

    const now = lasgumNow_();
    if (!record.createdAt) record.createdAt = now;
    record.updatedAt = now;
    record.deleted = record.deleted === true;
    record.deviceId = deviceId || record.deviceId || "";
    record.schemaVersion = LASGUM_SCHEMA_VERSION;
    record = normalizeRecord_(tableName, record);

    const row = headers.map(function(header) {
      return normalizeForSheet_(record[header], header, tableName);
    });

    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      headers.length
    ).setValues([row]);

    writeSyncLog_(
      deviceId,
      "CREATE",
      tableName,
      record.id,
      record.createdAt,
      now,
      "OK",
      "Record dibuat"
    );

    return {
      ok: true,
      action: "CREATE",
      table: tableName,
      data: record,
      serverTime: now
    };

  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   UPDATE
   ========================================================= */

function apiUpdate_(tableName, record, deviceId) {

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!record || !record.id) {
      throw new Error("UPDATE membutuhkan id.");
    }

    const sheet = lasgumGetSheet_(tableName);
    const headers = LASGUM_SHEETS[tableName];
    const rowNumber = findRowById_(sheet, headers, record.id);

    if (rowNumber === -1) {
      throw new Error("Record tidak ditemukan: " + record.id);
    }

    const merged = readRowObject_(sheet, headers, rowNumber);
    Object.keys(record).forEach(function(key) {
      if (headers.indexOf(key) !== -1) {
        merged[key] = record[key];
      }
    });

    const now = lasgumNow_();
    merged.updatedAt = now;
    merged.deviceId = deviceId || merged.deviceId || "";
    merged.schemaVersion = LASGUM_SCHEMA_VERSION;
    const normalized = normalizeRecord_(tableName, merged);

    const row = headers.map(function(header) {
      return normalizeForSheet_(normalized[header], header, tableName);
    });

    sheet.getRange(
      rowNumber,
      1,
      1,
      headers.length
    ).setValues([row]);

    writeSyncLog_(
      deviceId,
      "UPDATE",
      tableName,
      normalized.id,
      record.updatedAt || "",
      now,
      "OK",
      "Record diperbarui"
    );

    return {
      ok: true,
      action: "UPDATE",
      table: tableName,
      data: normalized,
      serverTime: now
    };

  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   SOFT DELETE
   ========================================================= */

function apiDelete_(tableName, id, deviceId) {

  if (!id) {
    throw new Error(
      "DELETE membutuhkan id."
    );
  }


  return apiUpdate_(
    tableName,
    {
      id: id,
      deleted: true,
      status: "deleted"
    },
    deviceId
  );

}


/* =========================================================
   SYNC
   ========================================================= */

function apiSync_(payload) {

  payload = payload || {};

  const deviceId =
    payload.deviceId || "";

  const records =
    Array.isArray(payload.records)
      ? payload.records
      : [];


  const results = [];

  records.forEach(function(item) {

    try {

      const tableName =
        item.tableName;

      const action =
        String(
          item.action || "CREATE"
        ).toUpperCase();

      const record =
        item.record || {};


      let result;


      if (action === "CREATE") {

        result =
          apiCreate_(
            tableName,
            record,
            deviceId
          );

      }

      else if (action === "UPDATE") {

        result =
          apiUpdate_(
            tableName,
            record,
            deviceId
          );

      }

      else if (action === "DELETE") {

        result =
          apiDelete_(
            tableName,
            record.id,
            deviceId
          );

      }

      else {

        throw new Error(
          "Action sync tidak dikenal: " +
          action
        );

      }


      results.push({

        ok: true,

        table:
          tableName,

        action:
          action,

        id:
          record.id,

        result:
          result

      });

    }

    catch (error) {

      results.push({

        ok: false,

        table:
          item.tableName || "",

        action:
          item.action || "",

        id:
          item.record &&
          item.record.id
            ? item.record.id
            : "",

        error:
          error.message

      });

    }

  });


  const failed = results.filter(function(item) { return !item.ok; });

  return {
    ok: failed.length === 0,
    partial: failed.length > 0 && failed.length < results.length,
    deviceId: deviceId,
    processed: results.length,
    succeeded: results.length - failed.length,
    failed: failed.length,
    results: results,
    serverTime: lasgumNow_(),
    schemaVersion: LASGUM_SCHEMA_VERSION
  };

}


/* =========================================================
   SYNC LOG
   ========================================================= */

function writeSyncLog_(
  deviceId,
  action,
  tableName,
  recordId,
  clientUpdatedAt,
  serverUpdatedAt,
  status,
  message
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("SYNC_LOG");

  if (!sheet) return;


  const row = [

    lasgumId_(),

    deviceId || "",

    action || "",

    tableName || "",

    recordId || "",

    clientUpdatedAt || "",

    serverUpdatedAt || "",

    status || "",

    message || "",

    lasgumNow_(),

    LASGUM_SCHEMA_VERSION

  ];


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      1,
      row.length
    )
    .setValues([row]);

}


/* =========================================================
   NILAI SHEET
   ========================================================= */

/* =========================================================
   NORMALISASI DATA
   ========================================================= */

const LASGUM_STRING_FIELDS = {
  SISWA: ["id","nis","nisn","nama","jenisKelamin","tempatLahir","tanggalLahir","kelas","rombel","fase","status","deviceId","ownerId","asalSekolah"],
  GURU: ["id","penggunaId","nik","nip","nuptk","nama","jenisKelamin","nomorHp","email","status","deviceId"],
  PENGGUNA: ["id","username","nama","role","passwordHash","status","deviceId"],
  MATA_PELAJARAN: ["id","kode","nama","kelas","fase","cakupan","scopeVersion","status","deviceId"],
  MATERI_CP: ["id","mataPelajaranId","kelas","fase","kodeCp","cp","material","deskripsi","semester","status","deviceId"],
  SOAL: ["id","mataPelajaranId","materiId","kelas","fase","tipe","pertanyaan","kunci","kataKunci","pembahasan","status","sumber","createdBy","deviceId"],
  PAKET_UJIAN: ["id","kode","nama","mataPelajaranId","kelas","fase","status","createdBy","deviceId"],
  PAKET_SOAL: ["id","paketId","soalId","deviceId"],
  JADWAL_UJIAN: ["id","paketId","aktif","status","deviceId"],
  UJIAN: ["id","paketId","siswaId","status","attempt","deviceId"],
  JAWABAN: ["id","ujianId","soalId","jawaban","manual","deviceId"],
  NILAI: ["id","ujianId","siswaId","paketId","status","catatan","deviceId"]
};

const LASGUM_DATE_ONLY_FIELDS = {
  SISWA: ["tanggalLahir"]
};

const LASGUM_DATETIME_FIELDS = {
  SEKOLAH: ["createdAt","updatedAt"],
  PENGGUNA: ["createdAt","updatedAt"],
  GURU: ["createdAt","updatedAt"],
  SISWA: ["createdAt","updatedAt"],
  MATA_PELAJARAN: ["createdAt","updatedAt"],
  MATERI_CP: ["createdAt","updatedAt"],
  SOAL: ["createdAt","updatedAt"],
  PAKET_UJIAN: ["createdAt","updatedAt"],
  PAKET_SOAL: ["createdAt","updatedAt"],
  JADWAL_UJIAN: ["mulai","selesai","createdAt","updatedAt"],
  UJIAN: ["mulai","selesai","createdAt","updatedAt"],
  JAWABAN: ["createdAt","updatedAt"],
  NILAI: ["createdAt","updatedAt"],
  PENGATURAN: ["updatedAt"]
};

const LASGUM_BOOLEAN_FIELDS = {
  SEKOLAH: ["deleted"],
  PENGGUNA: ["deleted"],
  GURU: ["deleted"],
  SISWA: ["deleted"],
  MATA_PELAJARAN: ["deleted"],
  MATERI_CP: ["deleted"],
  SOAL: ["deleted"],
  PAKET_UJIAN: ["acakSoal","acakOpsi","dipublikasikan","deleted"],
  PAKET_SOAL: ["deleted"],
  JADWAL_UJIAN: ["aktif","deleted"],
  UJIAN: ["deleted"],
  JAWABAN: ["benar","manual","deleted"],
  NILAI: ["deleted"]
};

function toBoolean_(value) {
  if (value === true || value === false) return value;
  const s = String(value == null ? "" : value).trim().toLowerCase();
  if (["true","1","ya","yes","aktif","active"].indexOf(s) !== -1) return true;
  if (["false","0","tidak","no","nonaktif","inactive",""] .indexOf(s) !== -1) return false;
  return Boolean(value);
}

function toDateOnly_(value) {
  if (value === null || value === undefined || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(value).trim();
  const m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[1] + "-" + String(m[2]).padStart(2,"0") + "-" + String(m[3]).padStart(2,"0");
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return s;
}

function toDateTime_(value) {
  if (value === null || value === undefined || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";
    return value.toISOString();
  }
  const d = new Date(String(value));
  return isNaN(d.getTime()) ? String(value) : d.toISOString();
}

function normalizeRecord_(tableName, record) {
  const out = {};
  record = record || {};
  Object.keys(record).forEach(function(key) {
    out[key] = record[key];
  });

  (LASGUM_STRING_FIELDS[tableName] || []).forEach(function(field) {
    if (out[field] !== undefined && out[field] !== null) out[field] = String(out[field]);
  });

  (LASGUM_DATE_ONLY_FIELDS[tableName] || []).forEach(function(field) {
    if (out[field] !== undefined) out[field] = toDateOnly_(out[field]);
  });

  (LASGUM_DATETIME_FIELDS[tableName] || []).forEach(function(field) {
    if (out[field] !== undefined && out[field] !== "") out[field] = toDateTime_(out[field]);
  });

  (LASGUM_BOOLEAN_FIELDS[tableName] || []).forEach(function(field) {
    if (out[field] !== undefined) out[field] = toBoolean_(out[field]);
  });

  /* Skor dasar selalu numerik dan aman. */
  if (tableName === "JAWABAN" && out.skor !== undefined && out.skor !== "") {
    const n = Number(out.skor);
    out.skor = isNaN(n) ? 0 : n;
  }
  if (tableName === "NILAI") {
    ["skorPg","skorIsian","skorUraian","nilaiAkhir"].forEach(function(field) {
      if (out[field] !== undefined && out[field] !== "") {
        const n = Number(out[field]);
        out[field] = isNaN(n) ? 0 : n;
      }
    });
  }
  if (tableName === "SOAL" && out.bobot !== undefined && out.bobot !== "") {
    const n = Number(out.bobot);
    out.bobot = isNaN(n) ? 1 : n;
  }
  if (tableName === "MATA_PELAJARAN" && out.scopeVersion !== undefined && out.scopeVersion !== "") {
    const n = Number(out.scopeVersion);
    out.scopeVersion = isNaN(n) ? 1 : n;
  }
  if (tableName === "PAKET_SOAL" && out.bobot !== undefined && out.bobot !== "") {
    const n = Number(out.bobot);
    out.bobot = isNaN(n) ? 1 : n;
  }

  return out;
}

function normalizeForSheet_(value, header, tableName) {
  if (value === undefined || value === null) return "";

  if ((LASGUM_DATE_ONLY_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return toDateOnly_(value);
  }

  if ((LASGUM_DATETIME_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return toDateTime_(value);
  }

  if (typeof value === "object") return JSON.stringify(value);
  return value;
}

function normalizeValue_(value, header, tableName) {
  if (value === undefined || value === null) return "";

  if ((LASGUM_DATE_ONLY_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return toDateOnly_(value);
  }

  if ((LASGUM_DATETIME_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return toDateTime_(value);
  }

  if ((LASGUM_STRING_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return String(value);
  }

  if ((LASGUM_BOOLEAN_FIELDS[tableName] || []).indexOf(header) !== -1) {
    return toBoolean_(value);
  }

  if (typeof value === "string" && (value.trim().startsWith("{") || value.trim().startsWith("["))) {
    try { return JSON.parse(value); } catch (err) { return value; }
  }

  return value;
}

function findRowById_(sheet, headers, id) {
  const idIndex = headers.indexOf("id");
  if (idIndex === -1 || !id) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function readRowObject_(sheet, headers, rowNumber) {
  const row = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = normalizeValue_(row[index], header, sheet.getName());
  });
  return obj;
}


/* =========================================================
   MIGRASI / KONSISTENSI SCHEMA V5
   ========================================================= */

/**
 * Menyamakan data lama ke format schema aktif tanpa menghapus data.
 * - schemaVersion -> versi schema aktif
 * - field teks -> string
 * - tanggalLahir SISWA -> YYYY-MM-DD
 * - datetime -> ISO
 * - boolean -> true/false
 * - angka skor/bobot -> number
 *
 * Jalankan sekali setelah memasang Code.gs V5 jika database lama perlu dimigrasikan.
 */
function migrateDatabaseToV5() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    ok: true,
    schemaVersion: LASGUM_SCHEMA_VERSION,
    database: ss.getName(),
    tables: {},
    serverTime: lasgumNow_()
  };

  Object.keys(LASGUM_SHEETS).forEach(function(tableName) {

    const sheet = ss.getSheetByName(tableName);
    const headers = LASGUM_SHEETS[tableName];

    if (!sheet || sheet.getLastRow() < 2) {
      result.tables[tableName] = {
        rows: 0,
        updated: 0
      };
      return;
    }

    const rowCount = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
    let updated = 0;

    const normalizedRows = values.map(function(row) {

      const record = {};
      headers.forEach(function(header, index) {
        record[header] = row[index];
      });

      const normalized = normalizeRecord_(tableName, record);
      normalized.schemaVersion = LASGUM_SCHEMA_VERSION;

      const outRow = headers.map(function(header) {
        return normalizeForSheet_(normalized[header], header, tableName);
      });

      updated++;
      return outRow;
    });

    sheet
      .getRange(2, 1, normalizedRows.length, headers.length)
      .setValues(normalizedRows);

    result.tables[tableName] = {
      rows: rowCount,
      updated: updated
    };
  });

  // Pastikan pengaturan schemaVersion ikut menjadi V5.
  const settings = ss.getSheetByName("PENGATURAN");
  if (settings && settings.getLastRow() >= 2) {
    const headers = LASGUM_SHEETS.PENGATURAN;
    const keyIndex = headers.indexOf("kunci");
    const valueIndex = headers.indexOf("nilai");
    const versionIndex = headers.indexOf("schemaVersion");

    if (keyIndex !== -1 && valueIndex !== -1) {
      const rowCount = settings.getLastRow() - 1;
      const values = settings.getRange(2, 1, rowCount, headers.length).getValues();

      values.forEach(function(row, i) {
        if (String(row[keyIndex]) === "schemaVersion") {
          row[valueIndex] = LASGUM_SCHEMA_VERSION;
          if (versionIndex !== -1) row[versionIndex] = LASGUM_SCHEMA_VERSION;
          settings.getRange(i + 2, 1, 1, headers.length).setValues([row]);
        }
      });
    }
  }

  SpreadsheetApp.flush();
  return result;
}

/**
 * Cek cepat bahwa data SISWA lama juga sudah keluar dalam format schema aktif.
 */
function testExistingStudentNormalizationV5() {

  const result = apiGet_("SISWA", {});
  if (!result.ok) {
    throw new Error("GET SISWA gagal.");
  }

  const invalid = [];

  result.data.forEach(function(row) {
    if (row.kelas !== "" && typeof row.kelas !== "string") {
      invalid.push(row.id + ": kelas bukan string");
    }

    if (row.rombel !== "" && typeof row.rombel !== "string") {
      invalid.push(row.id + ": rombel bukan string");
    }

    if (row.tanggalLahir !== "" &&
        !/^\d{4}-\d{2}-\d{2}$/.test(String(row.tanggalLahir))) {
      invalid.push(row.id + ": tanggalLahir bukan YYYY-MM-DD");
    }

    if (row.deleted !== true && row.deleted !== false) {
      invalid.push(row.id + ": deleted bukan boolean");
    }
  });

  if (invalid.length) {
    throw new Error(
      "Normalisasi SISWA belum konsisten:\n" + invalid.join("\n")
    );
  }

  Logger.log(JSON.stringify({
    ok: true,
    schemaVersion: LASGUM_SCHEMA_VERSION,
    siswaAktif: result.count,
    message: "Normalisasi data SISWA konsisten."
  }, null, 2));

  return {
    ok: true,
    test: "EXISTING_STUDENT_NORMALIZATION_V5",
    siswaAktif: result.count,
    schemaVersion: LASGUM_SCHEMA_VERSION
  };
}

/* =========================================================
   WEB APP GET
   ========================================================= */

function doGetLegacy_(e) {

  try {

    const params =
      e && e.parameter
        ? e.parameter
        : {};

    const action =
      String(
        params.action || "health"
      ).toLowerCase();


    if (action === "health") {

      return lasgumJson_(
        healthCheck()
      );

    }


    if (action === "meta") {

      return lasgumJson_({

        ok: true,

        database:
          SpreadsheetApp
            .getActiveSpreadsheet()
            .getName(),

        schemaVersion:
          LASGUM_SCHEMA_VERSION,

        serverTime:
          lasgumNow_(),

        tables:
          Object.keys(
            LASGUM_SHEETS
          )

      });

    }


    if (action === "get") {

      if (!params.table) {

        throw new Error(
          "Parameter table wajib diisi."
        );

      }


      return lasgumJson_(
        apiGet_(
          params.table,
          params
        )
      );

    }


    return lasgumJson_({

      ok: false,

      error:
        "Action GET tidak dikenal: " +
        action,

      serverTime:
        lasgumNow_()

    });

  }

  catch (error) {

    return lasgumJson_({

      ok: false,

      error:
        error.message,

      serverTime:
        lasgumNow_(),

      schemaVersion:
        LASGUM_SCHEMA_VERSION

    });

  }

}


/* =========================================================
   WEB APP POST
   ========================================================= */

function doPostLegacy_(e) {

  try {

    const body =
      lasgumParse_(e);

    const action =
      String(
        body.action || "health"
      ).toLowerCase();


    if (action === "health") {

      return lasgumJson_(
        healthCheck()
      );

    }


    if (action === "setup") {

      return lasgumJson_(
        setupDatabase()
      );

    }


    if (action === "create") {

      return lasgumJson_(
        apiCreate_(
          body.table,
          body.record,
          body.deviceId
        )
      );

    }


    if (action === "update") {

      return lasgumJson_(
        apiUpdate_(
          body.table,
          body.record,
          body.deviceId
        )
      );

    }


    if (action === "delete") {

      return lasgumJson_(
        apiDelete_(
          body.table,
          body.id,
          body.deviceId
        )
      );

    }


    if (action === "sync") {

      return lasgumJson_(
        apiSync_(body)
      );

    }


    return lasgumJson_({

      ok: false,

      error:
        "Action POST tidak dikenal: " +
        action,

      serverTime:
        lasgumNow_(),

      schemaVersion:
        LASGUM_SCHEMA_VERSION

    });

  }

  catch (error) {

    return lasgumJson_({

      ok: false,

      error:
        error.message,

      serverTime:
        lasgumNow_(),

      schemaVersion:
        LASGUM_SCHEMA_VERSION

    });

  }

}




/* =========================================================
   V32 SECURITY — DATABASE SERVER KHUSUS LASGUM
   Data API (GET table, CREATE, UPDATE, DELETE, SYNC, SETUP)
   wajib memakai token sesi LASGUM.
   Health/meta tetap publik untuk diagnosis koneksi.
   ========================================================= */
const LASGUM_API_SECURITY_VERSION = 1;
const LASGUM_API_TOKEN_TTL_SEC = 21600; // 6 jam
const LASGUM_SERVER_KEY_PROPERTY = "LASGUM_SERVER_KEY";
const LASGUM_TOKEN_PREFIX = "LASGUM_TOKEN_";

function getLasgumServerKey_() {
  return String(PropertiesService.getScriptProperties().getProperty(LASGUM_SERVER_KEY_PROPERTY) || "").trim();
}

function generateLasgumServerKey_() {
  const existing = getLasgumServerKey_();
  if (existing) return {ok:true, created:false, serverKey:existing, message:"LASGUM_SERVER_KEY sudah ada."};
  const key = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  PropertiesService.getScriptProperties().setProperty(LASGUM_SERVER_KEY_PROPERTY, key);
  Logger.log("LASGUM_SERVER_KEY: " + key);
  return {ok:true, created:true, serverKey:key, message:"Simpan kunci ini untuk dimasukkan ke Database Server LASGUM."};
}

function setupLasgumSecurity() {
  return generateLasgumServerKey_();
}

function issueLasgumToken_(deviceId) {
  const token = Utilities.getUuid() + "." + Utilities.getUuid();
  const payload = JSON.stringify({role:"lasgum", deviceId:String(deviceId || ""), issuedAt:Date.now()});
  CacheService.getScriptCache().put(LASGUM_TOKEN_PREFIX + token, payload, LASGUM_API_TOKEN_TTL_SEC);
  return token;
}

function requireLasgumToken_(token) {
  const raw = String(token || "").trim();
  if (!raw) throw new Error("AUTH_REQUIRED: Database Server hanya dapat diakses oleh LASGUM.");
  const item = CacheService.getScriptCache().get(LASGUM_TOKEN_PREFIX + raw);
  if (!item) throw new Error("AUTH_EXPIRED: Sesi Database LASGUM tidak valid atau sudah kedaluwarsa.");
  let data;
  try { data = JSON.parse(item); } catch(e) { throw new Error("AUTH_INVALID: Token Database LASGUM rusak."); }
  if (data.role !== "lasgum") throw new Error("AUTH_FORBIDDEN: Hak akses tidak mencukupi.");
  return data;
}

function authenticateLasgum_(serverKey, deviceId) {
  const configured = getLasgumServerKey_();
  if (!configured) throw new Error("SERVER_KEY_NOT_SET: Jalankan setupLasgumSecurity() pada Apps Script terlebih dahulu.");
  if (!serverKey || String(serverKey) !== configured) throw new Error("AUTH_FAILED: Kunci akses Database LASGUM salah.");
  const token = issueLasgumToken_(deviceId);
  return {ok:true, role:"lasgum", token:token, expiresInSeconds:LASGUM_API_TOKEN_TTL_SEC, serverTime:lasgumNow_(), securityVersion:LASGUM_API_SECURITY_VERSION};
}

function secureDoGet_(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || "health").toLowerCase();
  if (action === "health") return lasgumJson_(healthCheck());
  if (action === "meta") return lasgumJson_({ok:true,database:SpreadsheetApp.getActiveSpreadsheet().getName(),schemaVersion:LASGUM_SCHEMA_VERSION,securityVersion:LASGUM_API_SECURITY_VERSION,serverTime:lasgumNow_(),tables:Object.keys(LASGUM_SHEETS)});
  if (action === "get") {
    requireLasgumToken_(params.token || params.authToken);
    if (!params.table) throw new Error("Parameter table wajib diisi.");
    return lasgumJson_(apiGet_(params.table, params));
  }
  return lasgumJson_({ok:false,error:"Action GET tidak dikenal: " + action,serverTime:lasgumNow_(),schemaVersion:LASGUM_SCHEMA_VERSION});
}

function secureDoPost_(e) {
  const body = lasgumParse_(e);
  const action = String(body.action || "health").toLowerCase();
  if (action === "health") return lasgumJson_(healthCheck());
  if (action === "authenticate") return lasgumJson_(authenticateLasgum_(body.serverKey, body.deviceId));
  requireLasgumToken_(body.authToken || body.token);
  if (action === "setup") return lasgumJson_(setupDatabase());
  if (action === "create") return lasgumJson_(apiCreate_(body.table, body.record, body.deviceId));
  if (action === "update") return lasgumJson_(apiUpdate_(body.table, body.record, body.deviceId));
  if (action === "delete") return lasgumJson_(apiDelete_(body.table, body.id, body.deviceId));
  if (action === "sync") return lasgumJson_(apiSync_(body));
  return lasgumJson_({ok:false,error:"Action POST tidak dikenal: " + action,serverTime:lasgumNow_(),schemaVersion:LASGUM_SCHEMA_VERSION});
}

function doGet(e) {
  try { return secureDoGet_(e); }
  catch (error) { return lasgumJson_({ok:false,error:error.message,serverTime:lasgumNow_(),schemaVersion:LASGUM_SCHEMA_VERSION,securityVersion:LASGUM_API_SECURITY_VERSION}); }
}

function doPost(e) {
  try { return secureDoPost_(e); }
  catch (error) { return lasgumJson_({ok:false,error:error.message,serverTime:lasgumNow_(),schemaVersion:LASGUM_SCHEMA_VERSION,securityVersion:LASGUM_API_SECURITY_VERSION}); }
}

function testLasgumSecurity() {
  const key = getLasgumServerKey_();
  if (!key) throw new Error("LASGUM_SERVER_KEY belum dibuat. Jalankan setupLasgumSecurity().");
  const auth = authenticateLasgum_(key, "TEST_SECURITY_DEVICE");
  requireLasgumToken_(auth.token);
  Logger.log(JSON.stringify({ok:true,test:"LASGUM_DATABASE_SECURITY",securityVersion:LASGUM_API_SECURITY_VERSION,tokenValid:true,message:"Proteksi Database Server aktif."}, null, 2));
  return {ok:true,test:"LASGUM_DATABASE_SECURITY",securityVersion:LASGUM_API_SECURITY_VERSION,tokenValid:true};
}

function testSchemaV5() {
  const health = healthCheck();
  const meta = { ok: true, schemaVersion: LASGUM_SCHEMA_VERSION, tables: Object.keys(LASGUM_SHEETS) };
  if (health.schemaVersion !== LASGUM_SCHEMA_VERSION) throw new Error('Health/schemaVersion tidak konsisten.');
  if (meta.schemaVersion !== LASGUM_SCHEMA_VERSION || meta.tables.length !== 15) throw new Error('Meta/schema V5 tidak konsisten.');
  const mp = LASGUM_SHEETS.MATA_PELAJARAN;
  ['cakupan','scopeVersion'].forEach(function(field){
    if (mp.indexOf(field) === -1) throw new Error('MATA_PELAJARAN belum memiliki kolom '+field+'.');
  });
  Logger.log(JSON.stringify({ok:true,test:'SCHEMA_V5',schemaVersion:LASGUM_SCHEMA_VERSION,tableCount:meta.tables.length,message:'Schema V5 aktif dan 15 tabel terdaftar.'},null,2));
  return {ok:true,test:'SCHEMA_V5',schemaVersion:LASGUM_SCHEMA_VERSION,tableCount:meta.tables.length};
}

// Alias kompatibilitas untuk menu/test lama.
function testSchemaV3() { return testSchemaV5(); }


function testSiswaSchemaV4() {
  const headers = LASGUM_SHEETS.SISWA;
  const required = ["ownerId", "asalSekolah"];
  const missing = required.filter(function(x) { return headers.indexOf(x) === -1; });
  if (missing.length) throw new Error("Kolom SISWA belum lengkap: " + missing.join(", "));
  const sheet = lasgumGetSheet_("SISWA");
  const actual = sheet.getRange(1, 1, 1, headers.length).getValues()[0].map(String);
  const mismatch = headers.filter(function(h, i) { return actual[i] !== h; });
  if (mismatch.length) throw new Error("Header SISWA belum sesuai. Jalankan setupDatabase().");
  Logger.log(JSON.stringify({ok:true,test:"SISWA_SCHEMA_CURRENT",schemaVersion:LASGUM_SCHEMA_VERSION,columns:headers.length,message:"SISWA siap untuk integrasi aplikasi."},null,2));
  return {ok:true,test:"SISWA_SCHEMA_CURRENT",schemaVersion:LASGUM_SCHEMA_VERSION,columns:headers.length};
}

/* =========================================================
   TEST DATABASE
   ========================================================= */

function testDatabase() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const result = {

    ok: true,

    database:
      ss.getName(),

    schemaVersion:
      LASGUM_SCHEMA_VERSION,

    sheets: {}

  };


  Object.keys(
    LASGUM_SHEETS
  ).forEach(function(name) {

    const sheet =
      ss.getSheetByName(name);

    result.sheets[name] =
      sheet
        ? {
            exists: true,
            rows: sheet.getLastRow(),
            columns: sheet.getLastColumn()
          }
        : {
            exists: false
          };

  });


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}


/* =========================================================
   TEST API INTERNAL
   ========================================================= */

function testApiLocal() {

  const result = healthCheck();

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;

}

/* =========================================================
   TEST NORMALISASI V5
   ========================================================= */

function testNormalizationDatabase() {

  const deviceId = "TEST_NORMALISASI_V5";
  const id = lasgumId_();

  const input = {
    id: id,
    nis: 9001,
    nisn: 1234567890,
    nama: "SISWA TEST NORMALISASI",
    jenisKelamin: "L",
    tempatLahir: "TEST",
    tanggalLahir: "2015/04/07",
    kelas: 3,
    rombel: 1,
    fase: "B",
    status: "aktif"
  };

  const created = apiCreate_("SISWA", input, deviceId);
  const read = apiGet_("SISWA", { id: id });

  if (!created.ok || !read.ok || read.count !== 1) {
    throw new Error("TEST NORMALISASI gagal.");
  }

  const row = read.data[0];
  Logger.log(JSON.stringify({
    id: row.id,
    nis: row.nis,
    nisn: row.nisn,
    tanggalLahir: row.tanggalLahir,
    kelas: row.kelas,
    rombel: row.rombel,
    fase: row.fase
  }, null, 2));

  if (typeof row.nis !== "string" ||
      typeof row.nisn !== "string" ||
      typeof row.kelas !== "string" ||
      typeof row.rombel !== "string" ||
      row.tanggalLahir !== "2015-04-07") {
    throw new Error("Normalisasi tipe/format data belum benar.");
  }

  apiDelete_("SISWA", id, deviceId);

  Logger.log("=== TEST NORMALISASI V5 BERHASIL ===");
  return { ok: true, test: "NORMALISASI_V5", data: row };
}

/* =========================================================
   TEST IDEMPOTENT CREATE
   ========================================================= */

function testIdempotentCreate() {

  const deviceId = "TEST_IDEMPOTENT_V5";
  const id = lasgumId_();
  const record = {
    id: id,
    nis: "IDEMP001",
    nisn: "IDEMPNISN001",
    nama: "SISWA TEST IDEMPOTENT",
    jenisKelamin: "P",
    kelas: "2",
    rombel: "A",
    fase: "A",
    status: "aktif"
  };

  const first = apiCreate_("SISWA", record, deviceId);
  const second = apiCreate_("SISWA", record, deviceId);
  const read = apiGet_("SISWA", { id: id });

  if (!first.ok || !second.ok || !second.duplicate || read.count !== 1) {
    throw new Error("CREATE idempotent gagal: record berpotensi duplikat.");
  }

  apiDelete_("SISWA", id, deviceId);
  Logger.log("=== TEST IDEMPOTENT CREATE BERHASIL ===");
  return { ok: true, test: "IDEMPOTENT_CREATE", count: read.count };
}

