/************************************************************
 * LASGUM DATABASE API
 * VERSION 3.2 FINAL
 *
 * Offline First
 * Google Apps Script + Google Spreadsheet
 *
 * Schema Version : 6
 ************************************************************/

const LASGUM_SCHEMA_VERSION = 7;
const LASGUM_RELEASE = "7.1-MULTI-ADMIN-ACCESS-MANAGER";
const LASGUM_PUBLIC_BOOTSTRAP_VERSION = "7.0";


/* =========================================================
   NAMA DAN STRUKTUR DATABASE
   ========================================================= */

const LASGUM_SHEETS = {

  SEKOLAH: [
    "id","nama","npsn","alamat","desaKecamatan",
    "kabupatenProvinsi","kepalaSekolah","logoUrl",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  PENGGUNA: [
    "id","username","nama","role","passwordHash","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  GURU: [
    "id","ownerId","penggunaId","nik","nip","nuptk","nama",
    "jenisKelamin","nomorHp","email","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion"
  ],

  // Data siswa dibuat minimal untuk mengurangi kegagalan input/sinkronisasi.
  // Nama sekolah mengikuti identitas SEKOLAH aktif, bukan field per siswa.
  SISWA: [
    "id","nama","nis","nisn","kelas",
    "createdAt","updatedAt","deleted","deviceId","schemaVersion","ownerId"
  ],

  MATA_PELAJARAN: [
    "id","kode","nama","kelas","fase","cakupan","scopeVersion","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  MATERI_CP: [
    "id","mataPelajaranId","kelas","fase","kodeCp",
    "cp","material","deskripsi","semester","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  SOAL: [
    "id","mataPelajaranId","materiId","kelas","fase",
    "tipe","pertanyaan","opsi","kunci","kataKunci",
    "pembahasan","bobot","status","sumber","createdBy",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  PAKET_UJIAN: [
    "id","kode","nama","mataPelajaranId","kelas","fase",
    "durasiMenit","jumlahSoal","acakSoal","acakOpsi",
    "dipublikasikan","status","createdBy",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  PAKET_SOAL: [
    "id","paketId","soalId","nomor","bobot",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  JADWAL_UJIAN: [
    "id","paketId","mulai","selesai","aktif","status",
    "createdAt","updatedAt","deleted","deviceId",
    "schemaVersion","ownerId"
  ],

  UJIAN: [
    "id","paketId","siswaId","mulai","selesai",
    "status","attempt","deviceId","createdAt",
    "updatedAt","deleted","schemaVersion","ownerId"
  ],

  JAWABAN: [
    "id","ujianId","soalId","jawaban","benar","skor",
    "manual","createdAt","updatedAt","deleted",
    "deviceId","schemaVersion","ownerId"
  ],

  NILAI: [
    "id","ujianId","siswaId","paketId",
    "skorPg","skorIsian","skorUraian",
    "nilaiAkhir","status","catatan",
    "createdAt","updatedAt","deleted",
    "deviceId","schemaVersion","ownerId"
  ],

  PENGATURAN: [
    "id","kunci","nilai","deskripsi",
    "updatedAt","schemaVersion"
  ],

  SYNC_LOG: [
    "id","deviceId","action","tableName","recordId",
    "clientUpdatedAt","serverUpdatedAt","status",
    "message","createdAt","schemaVersion","ownerId"
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
  migrateLegacyOwnership_();
  SpreadsheetApp.flush();
  return {ok:true,schemaVersion:LASGUM_SCHEMA_VERSION,database:ss.getName(),sheets:Object.keys(LASGUM_SHEETS),serverTime:lasgumNow_()};
}

// Menetapkan ownerId pada data lama. Jika hanya ada satu Admin aktif, data lama
// otomatis menjadi milik Admin tersebut. Jika ada >1 Admin, data tanpa owner
// dibiarkan untuk penetapan eksplisit oleh LASGUM agar tidak salah pemilik.
function migrateLegacyOwnership_(){
  const p=lasgumGetSheet_("PENGGUNA"), ph=LASGUM_SHEETS.PENGGUNA, rows=readSheetObjects_(p,ph);
  const admins=rows.filter(function(r){const role=String(r.role||'').toLowerCase();const st=String(r.status||'aktif').toLowerCase();return !r.deleted && (role==='admin'||role==='guru') && st==='aktif';});
  if(admins.length!==1) return {assigned:0,skipped:true,reason:'Ambiguous owner: '+admins.length+' active admins'};
  const ownerId=String(admins[0].id||''); if(!ownerId) return {assigned:0};
  const tables=Object.keys(LASGUM_SHEETS).filter(function(t){return OWNER_SCOPED_TABLES_[t];});
  let assigned=0;
  tables.forEach(function(t){
    const sh=lasgumGetSheet_(t), headers=LASGUM_SHEETS[t], vals=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues():[];
    const oi=headers.indexOf('ownerId'); if(oi<0)return;
    let changed=false; vals.forEach(function(row){ if(String(row[oi]||'').trim()===''){row[oi]=ownerId;changed=true;assigned++;} });
    if(changed) sh.getRange(2,1,vals.length,headers.length).setValues(vals);
  });
  return {assigned:assigned,ownerId:ownerId};
}

function readSheetObjects_(sheet, headers){
  if(sheet.getLastRow()<2)return [];
  const vals=sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues();
  return vals.map(function(row){const o={};headers.forEach(function(h,i){o[h]=normalizeValue_(row[i],h,'')});return o;});
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
   MULTI-ADMIN OWNERSHIP
   ========================================================= */
const OWNER_SCOPED_TABLES_ = {
  SEKOLAH:true, PENGGUNA:true, GURU:true, SISWA:true, MATA_PELAJARAN:true,
  MATERI_CP:true, SOAL:true, PAKET_UJIAN:true, PAKET_SOAL:true, JADWAL_UJIAN:true,
  UJIAN:true, JAWABAN:true, NILAI:true, SYNC_LOG:true
};

function actorIsAdmin_(actor){ return !!actor && (String(actor.role||'').toLowerCase()==='admin' || String(actor.role||'').toLowerCase()==='guru') && String(actor.userId||'').trim()!==''; }
function actorIsLasgum_(actor){ return !!actor && String(actor.role||'').toLowerCase()==='lasgum'; }
function assertLasgumReadOnly_(actor, action){ if(actorIsLasgum_(actor)) throw new Error('LASGUM_ACCESS_FORBIDDEN: LASGUM tidak mengakses database operasional Admin melalui API umum.'); }
function assertOwnerAccess_(tableName, record, actor){
  if(!OWNER_SCOPED_TABLES_[tableName] || !actorIsAdmin_(actor)) return;
  const oid=String(actor.userId||'');
  const ro=String((record||{}).ownerId||'').trim();
  if(ro && ro!==oid) throw new Error('OWNER_FORBIDDEN: Data bukan milik Admin yang sedang login.');
}
function forceOwner_(tableName, record, actor){
  if(!OWNER_SCOPED_TABLES_[tableName] || !actorIsAdmin_(actor)) return record;
  const r=record||{}; r.ownerId=String(actor.userId||''); return r;
}

/* =========================================================
   GET DATA
   ========================================================= */

function apiGet_(tableName, params, actor) {

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

  let filtered = data.filter(function(row) {
    return String(row.deleted).toLowerCase() !== "true";
  });

  // Admin hanya menerima data miliknya. LASGUM menerima seluruh data untuk monitoring/rekap.
  if (actorIsAdmin_(actor) && OWNER_SCOPED_TABLES_[tableName]) {
    const oid=String(actor.userId||'');
    filtered=filtered.filter(function(row){ return String(row.ownerId||'')===oid; });
  }
  if (actorIsAdmin_(actor) && tableName==='PENGGUNA') {
    filtered=filtered.filter(function(row){ return String(row.id||'')===String(actor.userId||''); });
  }


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
   VALIDASI DATA SISWA SEDERHANA V6
   Hanya: nama, nis, nisn, kelas.
   ========================================================= */

function sanitizeStudentRecordV6_(record) {
  const r = record || {};
  return {
    id: String(r.id || lasgumId_()),
    nama: String(r.nama || "").trim(),
    nis: String(r.nis || "").trim(),
    nisn: String(r.nisn || "").trim(),
    kelas: String(r.kelas || "").trim(),
    createdAt: r.createdAt || "",
    updatedAt: r.updatedAt || "",
    deleted: r.deleted === true,
    deviceId: String(r.deviceId || "").trim(),
    schemaVersion: LASGUM_SCHEMA_VERSION
  };
}

function validateStudentRecordV6_(record) {
  const r = sanitizeStudentRecordV6_(record);
  if (!r.nama) throw new Error("SISWA_INVALID: Nama siswa wajib diisi.");
  if (!r.nis) throw new Error("SISWA_INVALID: NIS wajib diisi.");
  if (!r.kelas) throw new Error("SISWA_INVALID: Kelas wajib diisi.");
  return r;
}

/* =========================================================
   CREATE
   ========================================================= */

function apiCreate_(tableName, record, deviceId, actor) {
  assertLasgumReadOnly_(actor, "CREATE");
  record = forceOwner_(String(tableName||'').toUpperCase(), record || {}, actor);
  assertOwnerAccess_(String(tableName||'').toUpperCase(), record, actor);

  const lock = LockService.getScriptLock();
  if (String(tableName).toUpperCase() === "SISWA") {
    record = validateStudentRecordV6_(record);
    record = forceOwner_(String(tableName||'').toUpperCase(), record, actor);
  }
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

function apiUpdate_(tableName, record, deviceId, actor) {
  assertLasgumReadOnly_(actor, "UPDATE");

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

    let merged = readRowObject_(sheet, headers, rowNumber);
    assertOwnerAccess_(String(tableName||'').toUpperCase(), merged, actor);
    Object.keys(record).forEach(function(key) {
      if (headers.indexOf(key) !== -1) {
        merged[key] = record[key];
      }
    });

    if (String(tableName).toUpperCase() === "SISWA") {
      if (merged.deleted === true) {
        // Soft delete tetap membutuhkan record lama yang sudah ada.
        merged = sanitizeStudentRecordV6_(merged);
      } else {
        merged = validateStudentRecordV6_(merged);
      }
    }

    merged = forceOwner_(String(tableName||'').toUpperCase(), merged, actor);
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

function apiDelete_(tableName, id, deviceId, actor) {
  assertLasgumReadOnly_(actor, "DELETE");

  if (!id) {
    throw new Error(
      "DELETE membutuhkan id."
    );
  }


  const sh=lasgumGetSheet_(tableName), hs=LASGUM_SHEETS[tableName], rn=findRowById_(sh,hs,id);
  if(rn===-1) throw new Error('Record tidak ditemukan: '+id);
  const existing=readRowObject_(sh,hs,rn);
  assertOwnerAccess_(String(tableName||'').toUpperCase(), existing, actor);
  return apiUpdate_(tableName,{id:id,deleted:true,status:"deleted"},deviceId,actor);

}


/* =========================================================
   SYNC
   ========================================================= */

function apiSync_(payload, actor) {
  assertLasgumReadOnly_(actor, "SYNC");

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
            deviceId,
            actor
          );

      }

      else if (action === "UPDATE") {

        result =
          apiUpdate_(
            tableName,
            record,
            deviceId,
            actor
          );

      }

      else if (action === "DELETE") {

        result =
          apiDelete_(
            tableName,
            record.id,
            deviceId,
            actor
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
  SISWA: ["id","nama","nis","nisn","kelas","deviceId"],
  GURU: ["id","penggunaId","nik","nip","nuptk","nama","jenisKelamin","nomorHp","email","status","deviceId"],
  PENGGUNA: ["id","username","nama","role","passwordHash","status","deviceId"],
  MATA_PELAJARAN: ["id","kode","nama","kelas","fase","cakupan","scopeVersion","status","deviceId"],
  MATERI_CP: ["id","mataPelajaranId","kelas","fase","kodeCp","cp","material","deskripsi","semester","status","deviceId"],
  SOAL: ["id","mataPelajaranId","materiId","kelas","fase","tipe","pertanyaan","kunci","kataKunci","pembahasan","status","sumber","createdBy","deviceId"],
  PAKET_UJIAN: ["id","kode","nama","mataPelajaranId","kelas","fase","status","createdBy","deviceId"],
  PAKET_SOAL: ["id","paketId","soalId","deviceId"],
  JADWAL_UJIAN: ["id","paketId","aktif","status","deviceId","ownerId"],
  UJIAN: ["id","paketId","siswaId","status","attempt","deviceId"],
  JAWABAN: ["id","ujianId","soalId","jawaban","manual","deviceId","ownerId"],
  NILAI: ["id","ujianId","siswaId","paketId","status","catatan","deviceId","ownerId"]
};

const LASGUM_DATE_ONLY_FIELDS = {};

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
   MIGRASI / KONSISTENSI SCHEMA V6
   ========================================================= */

/**
 * Menyamakan data lama ke format schema aktif tanpa menghapus data.
 * - schemaVersion -> versi schema aktif
 * - field teks -> string
 * - datetime -> ISO
 * - boolean -> true/false
 * - angka skor/bobot -> number
 *
 * Jalankan sekali setelah memasang Code.gs V6 jika database lama perlu dimigrasikan.
 */
function migrateDatabaseToV6() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SISWA");

  if (!sheet) {
    setupDatabase();
    return { ok: true, created: true, migrated: 0, schemaVersion: LASGUM_SCHEMA_VERSION };
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const oldHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const oldSchema = oldHeaders.indexOf("jenisKelamin") >= 0 ||
                    oldHeaders.indexOf("asalSekolah") >= 0 ||
                    oldHeaders.indexOf("rombel") >= 0;

  if (!oldSchema) {
    setupDatabase();
    return { ok: true, migrated: 0, message: "SISWA sudah menggunakan schema sederhana V6." };
  }

  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, oldHeaders.length).getValues()
    : [];

  function valueOf(row, key) {
    const i = oldHeaders.indexOf(key);
    return i >= 0 ? row[i] : "";
  }

  const newRows = rows.map(function(row) {
    const now = lasgumNow_();
    return [
      String(valueOf(row, "id") || lasgumId_()),
      String(valueOf(row, "nama") || "").trim(),
      String(valueOf(row, "nis") || "").trim(),
      String(valueOf(row, "nisn") || "").trim(),
      String(valueOf(row, "kelas") || "").trim(),
      valueOf(row, "createdAt") || now,
      valueOf(row, "updatedAt") || now,
      valueOf(row, "deleted") === true,
      String(valueOf(row, "deviceId") || ""),
      LASGUM_SCHEMA_VERSION
    ];
  });

  const headers = LASGUM_SHEETS.SISWA;
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    migrated: newRows.length,
    schemaVersion: LASGUM_SCHEMA_VERSION,
    message: "SISWA berhasil dimigrasikan ke schema sederhana V6."
  };
}

function migrateDatabaseToV5() {
  return migrateDatabaseToV6();
}

function testExistingStudentNormalizationV6() {
  const result = apiGet_("SISWA", {});
  if (!result.ok) throw new Error("GET SISWA gagal.");

  const invalid = (result.data || []).filter(function(row) {
    return !("id" in row) || !("nama" in row) || !("nis" in row) ||
           !("nisn" in row) || !("kelas" in row);
  });

  if (invalid.length) {
    throw new Error("Masih ada record SISWA dengan schema lama.");
  }

  return {
    ok: true,
    test: "EXISTING_STUDENT_NORMALIZATION_V6",
    siswa: result.count,
    schemaVersion: LASGUM_SCHEMA_VERSION
  };
}

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
const LASGUM_USER_TOKEN_PREFIX = "LASGUM_USER_TOKEN_";

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

function issueLasgumUserToken_(userId, username, deviceId, role) {
  const token = Utilities.getUuid() + "." + Utilities.getUuid();
  const payload = JSON.stringify({role:String(role||"guru").toLowerCase(), userId:String(userId||""), username:String(username||""), deviceId:String(deviceId||""), issuedAt:Date.now()});
  CacheService.getScriptCache().put(LASGUM_USER_TOKEN_PREFIX + token, payload, LASGUM_API_TOKEN_TTL_SEC);
  return token;
}

function requireLasgumActor_(token){ return requireLasgumToken_(token); }

function lasgumCreateAdmin_(body, actor){
  requireLasgumActor_(body.authToken||body.token);
  const username=String(body.username||'').trim();
  const nama=String(body.nama||'').trim();
  const sekolah=String(body.sekolah||'').trim();
  const passwordHash=String(body.passwordHash||'').trim();
  if(!username||!nama||!passwordHash) throw new Error('ADMIN_CREATE_FAILED: Nama, username, dan password wajib diisi.');
  const ps=readSheetObjects_(lasgumGetSheet_('PENGGUNA'),LASGUM_SHEETS.PENGGUNA);
  if(ps.some(x=>String(x.username||'').trim().toLowerCase()===username.toLowerCase()&&!x.deleted)) throw new Error('ADMIN_CREATE_FAILED: Username sudah digunakan.');
  const id=String(body.id||lasgumId_()), now=lasgumNow_();
  const p={id:id,username:username,nama:nama,role:'admin',passwordHash:passwordHash,status:'aktif',createdAt:now,updatedAt:now,deleted:false,deviceId:String(body.deviceId||''),schemaVersion:LASGUM_SCHEMA_VERSION,ownerId:id};
  apiCreate_('PENGGUNA',p,body.deviceId||'',{role:'admin',userId:id});
  const g={id:id,penggunaId:id,nama:nama,status:'aktif',createdAt:now,updatedAt:now,deleted:false,deviceId:String(body.deviceId||''),schemaVersion:LASGUM_SCHEMA_VERSION,ownerId:id};
  apiCreate_('GURU',g,body.deviceId||'',{role:'admin',userId:id});
  if(sekolah){
    const sh={id:lasgumId_(),nama:sekolah,npsn:'',alamat:'',desaKecamatan:'',kabupatenProvinsi:'',kepalaSekolah:'',logoUrl:'',createdAt:now,updatedAt:now,deleted:false,deviceId:String(body.deviceId||''),schemaVersion:LASGUM_SCHEMA_VERSION,ownerId:id};
    apiCreate_('SEKOLAH',sh,body.deviceId||'',{role:'admin',userId:id});
  }
  return {ok:true,action:'ADMIN_CREATE',data:{id:id,username:username,nama:nama,sekolah:sekolah,role:'admin',status:'aktif'},serverTime:now};
}

function lasgumSetAdminStatus_(body, actor){
  requireLasgumActor_(body.authToken||body.token);
  const id=String(body.userId||body.id||'').trim(); if(!id) throw new Error('ADMIN_STATUS_FAILED: userId wajib diisi.');
  const sh=lasgumGetSheet_('PENGGUNA'), hs=LASGUM_SHEETS.PENGGUNA, rn=findRowById_(sh,hs,id); if(rn===-1) throw new Error('ADMIN_STATUS_FAILED: Admin tidak ditemukan.');
  const cur=readRowObject_(sh,hs,rn); if(cur.deleted===true) throw new Error('ADMIN_STATUS_FAILED: Akun sudah dihapus.');
  const st=String(body.status||'aktif').toLowerCase()==='nonaktif'?'nonaktif':'aktif';
  apiUpdate_('PENGGUNA',{id:id,status:st},body.deviceId||'',{role:'admin',userId:id});
  const gh=LASGUM_SHEETS.GURU, gs=lasgumGetSheet_('GURU'), gr=findRowById_(gs,gh,id); if(gr!==-1) apiUpdate_('GURU',{id:id,status:st},body.deviceId||'',{role:'admin',userId:id});
  return {ok:true,action:'ADMIN_STATUS',data:{id:id,status:st},serverTime:lasgumNow_()};
}

function lasgumDeleteAdminCascade_(body, actor){
  requireLasgumActor_(body.authToken||body.token);
  const id=String(body.userId||body.id||'').trim(); if(!id) throw new Error('ADMIN_DELETE_FAILED: userId wajib diisi.');
  const p=readSheetObjects_(lasgumGetSheet_('PENGGUNA'),LASGUM_SHEETS.PENGGUNA).find(x=>String(x.id||'')===id);
  if(!p) throw new Error('ADMIN_DELETE_FAILED: Admin tidak ditemukan.');
  const tables=Object.keys(LASGUM_SHEETS).filter(t=>OWNER_SCOPED_TABLES_[t]);
  let deleted=0;
  tables.forEach(function(t){
    const sh=lasgumGetSheet_(t), hs=LASGUM_SHEETS[t], oi=hs.indexOf('ownerId');
    if(oi<0||sh.getLastRow()<2)return;
    const vals=sh.getRange(2,1,sh.getLastRow()-1,hs.length).getValues();
    const keep=[]; vals.forEach(function(row){ if(String(row[oi]||'').trim()===id){deleted++;} else keep.push(row); });
    if(keep.length!==vals.length){
      sh.getRange(2,1,vals.length,hs.length).clearContent();
      if(keep.length) sh.getRange(2,1,keep.length,hs.length).setValues(keep);
    }
  });
  return {ok:true,action:'ADMIN_DELETE_CASCADE',userId:id,deletedRecords:deleted,serverTime:lasgumNow_()};
}

function authenticateUser_(username, passwordHash, deviceId) {
  const u = String(username || "").trim().toLowerCase();
  const ph = String(passwordHash || "").trim();
  if (!u || !ph) throw new Error("AUTH_FAILED: Username dan password wajib diisi.");
  const sheet = lasgumGetSheet_("PENGGUNA");
  const headers = LASGUM_SHEETS.PENGGUNA;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("AUTH_FAILED: Belum ada akun Admin/Guru di Database Server.");
  const values = sheet.getRange(2,1,lastRow-1,headers.length).getValues();
  for (let i=0;i<values.length;i++) {
    const row={}; headers.forEach((h,j)=>row[h]=normalizeValue_(values[i][j],h,"PENGGUNA"));
    const role=String(row.role||"").toLowerCase();
    const status=String(row.status||"aktif").toLowerCase();
    if (String(row.username||"").trim().toLowerCase()===u && String(row.passwordHash||"")===ph) {
      if (row.deleted===true || status==="nonaktif" || status==="deleted") throw new Error("AUTH_FORBIDDEN: Akun tidak aktif.");
      const effectiveRole=(role==='guru'||role==='admin')?'admin':role;
      const token=issueLasgumUserToken_(row.id,row.username,deviceId,effectiveRole);
      return {ok:true,role:effectiveRole,userId:String(row.id),username:String(row.username),token:token,expiresInSeconds:LASGUM_API_TOKEN_TTL_SEC,serverTime:lasgumNow_(),securityVersion:LASGUM_API_SECURITY_VERSION};
    }
  }
  throw new Error("AUTH_FAILED: Username atau password Admin/Guru salah.");
}


/* =========================================================
   V6.2 — TOKEN SISWA + WRITE TERBATAS
   Siswa boleh mengirim hasil ujian ke server, tetapi hanya untuk
   record miliknya sendiri. Siswa tidak boleh menulis master database.
   ========================================================= */
const LASGUM_STUDENT_TOKEN_PREFIX = "LASGUM_STUDENT_TOKEN_";

function issueStudentToken_(studentId, nis, deviceId, ownerId) {
  const token = Utilities.getUuid() + "." + Utilities.getUuid();
  const payload = JSON.stringify({
    role: "siswa",
    studentId: String(studentId || ""),
    nis: String(nis || ""),
    deviceId: String(deviceId || ""),
    ownerId: String(ownerId || ""),
    issuedAt: Date.now()
  });
  CacheService.getScriptCache().put(LASGUM_STUDENT_TOKEN_PREFIX + token, payload, LASGUM_API_TOKEN_TTL_SEC);
  return token;
}

function authenticateStudent_(studentId, nis, deviceId) {
  const id = String(studentId || "").trim();
  const n = String(nis || "").trim();
  if (!id || !n) throw new Error("STUDENT_AUTH_FAILED: ID siswa dan NIS wajib diisi.");
  const result = apiGet_("SISWA", {});
  const rows = Array.isArray(result.data) ? result.data : [];
  const s = rows.find(function(row) {
    return row.deleted !== true && String(row.id || "") === id && String(row.nis || "").trim() === n;
  });
  if (!s) throw new Error("STUDENT_AUTH_FAILED: Data siswa tidak ditemukan atau NIS tidak cocok.");
  const token = issueStudentToken_(id, n, deviceId, s.ownerId);
  return {
    ok: true,
    role: "siswa",
    studentId: id,
    nis: n,
    token: token,
    expiresInSeconds: LASGUM_API_TOKEN_TTL_SEC,
    serverTime: lasgumNow_(),
    securityVersion: LASGUM_API_SECURITY_VERSION
  };
}

function requireStudentToken_(token) {
  const raw = String(token || "").trim();
  if (!raw) throw new Error("STUDENT_AUTH_REQUIRED: Sesi siswa belum aktif.");
  const item = CacheService.getScriptCache().get(LASGUM_STUDENT_TOKEN_PREFIX + raw);
  if (!item) throw new Error("STUDENT_AUTH_EXPIRED: Sesi siswa tidak valid atau sudah kedaluwarsa.");
  let data;
  try { data = JSON.parse(item); } catch (e) { throw new Error("STUDENT_AUTH_INVALID: Token siswa rusak."); }
  if (data.role !== "siswa" || !data.studentId) throw new Error("STUDENT_AUTH_FORBIDDEN: Hak akses siswa tidak valid.");
  return data;
}

function apiStudentSync_(payload, actor) {
  payload = payload || {};
  const records = Array.isArray(payload.records) ? payload.records : [];
  if (!records.length) return {ok:true,role:"siswa",processed:0,succeeded:0,failed:0,results:[],serverTime:lasgumNow_()};

  const allowedTables = { UJIAN:true, JAWABAN:true, NILAI:true };
  const allowedExamIds = {};
  const existingExams = {};
  const ex = apiGet_("UJIAN", {}).data || [];
  ex.forEach(function(row){ if (String(row.siswaId || "") === String(actor.studentId)) existingExams[String(row.id || "")] = true; });

  records.forEach(function(item){
    const t = String(item.tableName || "").toUpperCase();
    const r = item.record || {};
    if (t === "UJIAN" && String(r.siswaId || "") === String(actor.studentId)) allowedExamIds[String(r.id || "")] = true;
  });
  Object.keys(existingExams).forEach(function(id){ allowedExamIds[id] = true; });

  const results = [];
  records.forEach(function(item){
    try {
      const table = String(item.tableName || "").toUpperCase();
      const action = String(item.action || "CREATE").toUpperCase();
      const record = item.record || {};
      if (!allowedTables[table]) throw new Error("STUDENT_WRITE_FORBIDDEN: Siswa hanya boleh mengirim UJIAN, JAWABAN, dan NILAI.");
      if (action !== "CREATE" && action !== "UPDATE") throw new Error("STUDENT_WRITE_FORBIDDEN: Siswa tidak boleh menghapus data ujian.");

      if (table === "UJIAN" || table === "NILAI") {
        if (String(record.siswaId || "") !== String(actor.studentId)) throw new Error("STUDENT_WRITE_FORBIDDEN: siswaId tidak sesuai dengan sesi siswa.");
      }
      if (table === "JAWABAN") {
        const examId = String(record.ujianId || "");
        if (!examId || !allowedExamIds[examId]) throw new Error("STUDENT_WRITE_FORBIDDEN: Ujian bukan milik siswa pada sesi ini.");
      }

      if (actor.ownerId && OWNER_SCOPED_TABLES_[table]) record.ownerId=String(actor.ownerId);
      let out;
      if (action === "CREATE") out = apiCreate_(table, record, actor.deviceId || payload.deviceId || "", actor);
      else out = apiUpdate_(table, record, actor.deviceId || payload.deviceId || "", actor);
      results.push({ok:true,table:table,action:action,id:record.id || "",data:out});
    } catch (err) {
      results.push({ok:false,table:String(item.tableName || ""),action:String(item.action || "CREATE"),error:String(err && err.message || err)});
    }
  });
  const failed = results.filter(function(x){return !x.ok;}).length;
  return {ok:failed===0,role:"siswa",processed:results.length,succeeded:results.length-failed,failed:failed,results:results,serverTime:lasgumNow_()};
}

function requireAnyDatabaseToken_(token) {
  const raw=String(token||"").trim();
  if(!raw) throw new Error("LOGIN_LASGUM_REQUIRED: Fitur pengelolaan Database Pusat memerlukan login LASGUM.");
  const c=CacheService.getScriptCache();
  let item=c.get(LASGUM_TOKEN_PREFIX+raw);
  if(item){ let data=JSON.parse(item); if(data.role!=="lasgum") throw new Error("AUTH_FORBIDDEN: Hak akses tidak mencukupi."); return data; }
  item=c.get(LASGUM_USER_TOKEN_PREFIX+raw);
  if(item){ let data=JSON.parse(item); if(data.role!=="admin" && data.role!=="guru") throw new Error("AUTH_FORBIDDEN: Hak akses tidak mencukupi."); return data; }
  throw new Error("AUTH_EXPIRED: Sesi Database tidak valid atau sudah kedaluwarsa.");
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

const LASGUM_USERNAME_PROPERTY = "LASGUM_USERNAME";
const LASGUM_PASSWORD_HASH_PROPERTY = "LASGUM_PASSWORD_HASH";

function getLasgumUsername_(){
  return String(PropertiesService.getScriptProperties().getProperty(LASGUM_USERNAME_PROPERTY) || "lasgum").trim();
}
function getLasgumPasswordHash_(){
  return String(PropertiesService.getScriptProperties().getProperty(LASGUM_PASSWORD_HASH_PROPERTY) || "").trim().toLowerCase();
}
function setupLasgumAccount(){
  // Akun awal kompatibilitas V57. Setelah deployment, ganti password melalui fungsi ini.
  PropertiesService.getScriptProperties().setProperties({
    LASGUM_USERNAME:"lasgum",
    LASGUM_PASSWORD_HASH:"58b7fb952cf7fa769b4962e3f4077d77a1a9b983b379442ac8a21f2a9a97fb78"
  }, true);
  return {ok:true,username:"lasgum",message:"Akun LASGUM awal tersimpan di Script Properties. Ganti password setelah deployment."};
}
function setLasgumAccount(username,passwordHash){
  username=String(username||'').trim(); passwordHash=String(passwordHash||'').trim().toLowerCase();
  if(!username||!passwordHash) throw new Error('LASGUM_ACCOUNT_INVALID: username dan passwordHash wajib diisi.');
  PropertiesService.getScriptProperties().setProperties({LASGUM_USERNAME:username,LASGUM_PASSWORD_HASH:passwordHash}, true);
  return {ok:true,username:username,message:'Akun LASGUM diperbarui.'};
}
function authenticateLasgumUser_(username,passwordHash,deviceId){
  const u=getLasgumUsername_(), h=getLasgumPasswordHash_();
  if(!h) throw new Error('LASGUM_ACCOUNT_NOT_SET: Jalankan setupLasgumAccount() pada Apps Script terlebih dahulu.');
  if(String(username||'').trim()!==u || String(passwordHash||'').trim().toLowerCase()!==h) throw new Error('AUTH_FAILED: Username atau password LASGUM salah.');
  const token=issueLasgumToken_(deviceId);
  return {ok:true,role:'lasgum',username:u,token:token,expiresInSeconds:LASGUM_API_TOKEN_TTL_SEC,serverTime:lasgumNow_(),securityVersion:LASGUM_API_SECURITY_VERSION};
}
function lasgumListAdmins_(actor){
  if(!actorIsLasgum_(actor)) throw new Error('AUTH_FORBIDDEN: Hanya LASGUM yang dapat mengelola daftar Admin.');
  const rows=readSheetObjects_(lasgumGetSheet_('PENGGUNA'),LASGUM_SHEETS.PENGGUNA);
  const admins=rows.filter(function(r){return !r.deleted && ['admin','guru'].indexOf(String(r.role||'').toLowerCase())>=0;}).map(function(r){
    return {id:String(r.id||''),username:String(r.username||''),nama:String(r.nama||''),role:'admin',status:String(r.status||'aktif').toLowerCase()==='nonaktif'?'nonaktif':'aktif',active:String(r.status||'aktif').toLowerCase()!=='nonaktif',createdAt:r.createdAt||'',updatedAt:r.updatedAt||''};
  });
  const sekolah=readSheetObjects_(lasgumGetSheet_('SEKOLAH'),LASGUM_SHEETS.SEKOLAH);
  const schoolByOwner={}; sekolah.forEach(function(r){if(r.ownerId)schoolByOwner[String(r.ownerId)]=String(r.nama||'');});
  admins.forEach(function(a){a.sekolah=schoolByOwner[a.id]||'';});
  return {ok:true,action:'LASGUM_LIST_ADMINS',data:admins,counts:{admins:admins.length},serverTime:lasgumNow_()};
}

function authenticateLasgum_(serverKey, deviceId) {
  const configured = getLasgumServerKey_();
  if (!configured) throw new Error("SERVER_KEY_NOT_SET: Jalankan setupLasgumSecurity() pada Apps Script terlebih dahulu.");
  if (!serverKey || String(serverKey) !== configured) throw new Error("AUTH_FAILED: Kunci akses Database LASGUM salah.");
  const token = issueLasgumToken_(deviceId);
  return {ok:true, role:"lasgum", token:token, expiresInSeconds:LASGUM_API_TOKEN_TTL_SEC, serverTime:lasgumNow_(), securityVersion:LASGUM_API_SECURITY_VERSION};
}


/* ============================================================
   V46 PUBLIC DATABASE BOOTSTRAP
   Database read-only untuk semua perangkat.
   Tidak membutuhkan login LASGUM.
   PasswordHash PENGGUNA sengaja tidak dikirim ke browser.
   ============================================================ */

/**
 * V6 CENTRAL: bootstrap ringan khusus portal siswa.
 * Hanya mengirim identitas sekolah + siswa aktif.
 * Tidak membutuhkan token.
 */
function publicStudentsBootstrap_() {
  const sekolahResult = apiGet_("SEKOLAH", {}, null);
  const siswaResult = apiGet_("SISWA", {}, null);

  const sekolah = (sekolahResult.data || []).filter(function(s) {
    return s.deleted !== true &&
      String(s.status || "aktif").toLowerCase() !== "nonaktif";
  });

  const siswa = (siswaResult.data || []).filter(function(s) {
    return s.deleted !== true;
  }).map(function(s) {
    return {
      id: String(s.id || ""),
      nama: String(s.nama || ""),
      nis: String(s.nis || ""),
      nisn: String(s.nisn || ""),
      kelas: String(s.kelas || "")
    };
  });

  return {
    ok: true,
    mode: "public-students-bootstrap-v6.1",
    release: LASGUM_RELEASE,
    bootstrapVersion: LASGUM_PUBLIC_BOOTSTRAP_VERSION,
    schemaVersion: LASGUM_SCHEMA_VERSION,
    serverTime: lasgumNow_(),
    data: { SEKOLAH: sekolah, SISWA: siswa },
    counts: { sekolah: sekolah.length, siswa: siswa.length }
  };
}

function publicBootstrapV6_() {
  const school=publicStudentsBootstrap_();
  return {ok:true,mode:"public-safe-bootstrap-v7",release:LASGUM_RELEASE,bootstrapVersion:LASGUM_PUBLIC_BOOTSTRAP_VERSION,schemaVersion:LASGUM_SCHEMA_VERSION,serverTime:lasgumNow_(),data:{SEKOLAH:school.data.SEKOLAH,SISWA:school.data.SISWA},counts:{SEKOLAH:school.counts.sekolah,SISWA:school.counts.siswa}};
}

function publicBootstrap_() { return publicBootstrapV6_(); }

function secureDoGet_(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || "health").toLowerCase();
  if (action === "health") return lasgumJson_(healthCheck());
  if (action === "meta") return lasgumJson_({ok:true,database:SpreadsheetApp.getActiveSpreadsheet().getName(),release:LASGUM_RELEASE,bootstrapVersion:LASGUM_PUBLIC_BOOTSTRAP_VERSION,schemaVersion:LASGUM_SCHEMA_VERSION,securityVersion:LASGUM_API_SECURITY_VERSION,serverTime:lasgumNow_(),tables:Object.keys(LASGUM_SHEETS)});
  if (action === "bootstrap") return lasgumJson_(publicBootstrapV6_());
  if (action === "students") return lasgumJson_(publicStudentsBootstrap_());
  if (action === "get") {
    const actor=requireAnyDatabaseToken_(params.token || params.authToken);
    if (!params.table) throw new Error("Parameter table wajib diisi.");
    return lasgumJson_(apiGet_(params.table, params, actor));
  }
  return lasgumJson_({ok:false,error:"Action GET tidak dikenal: " + action,serverTime:lasgumNow_(),schemaVersion:LASGUM_SCHEMA_VERSION});
}

function secureDoPost_(e) {
  const body = lasgumParse_(e);
  const action = String(body.action || "health").toLowerCase();
  if (action === "health") return lasgumJson_(healthCheck());
  if (action === "authenticate") return lasgumJson_(authenticateLasgum_(body.serverKey, body.deviceId));
  if (action === "authenticateLasgumUser") return lasgumJson_(authenticateLasgumUser_(body.username, body.passwordHash, body.deviceId));
  if (action === "setupLasgumAccount") return lasgumJson_(setupLasgumAccount());
  if (action === "setLasgumAccount") return lasgumJson_(setLasgumAccount(body.username, body.passwordHash));
  if (action === "authenticateUser") return lasgumJson_(authenticateUser_(body.username, body.passwordHash, body.deviceId));
  if (action === "authenticateStudent") return lasgumJson_(authenticateStudent_(body.studentId, body.nis, body.deviceId));

  const rawToken = body.authToken || body.token;
  if (action === "lasgumListAdmins") return lasgumJson_(lasgumListAdmins_(requireLasgumActor_(rawToken)));
  if (action === "lasgumCreateAdmin") return lasgumJson_(lasgumCreateAdmin_(body, requireLasgumActor_(rawToken)));
  if (action === "lasgumSetAdminStatus") return lasgumJson_(lasgumSetAdminStatus_(body, requireLasgumActor_(rawToken)));
  if (action === "lasgumDeleteAdmin") return lasgumJson_(lasgumDeleteAdminCascade_(body, requireLasgumActor_(rawToken)));
  const cache = CacheService.getScriptCache();
  if (rawToken && cache.get(LASGUM_STUDENT_TOKEN_PREFIX + String(rawToken))) {
    const actor = requireStudentToken_(rawToken);
    if (action === "sync") return lasgumJson_(apiStudentSync_(body, actor));
    throw new Error("STUDENT_WRITE_FORBIDDEN: Sesi siswa hanya dapat digunakan untuk sinkronisasi hasil ujian.");
  }
  const actor=requireAnyDatabaseToken_(rawToken);
  if (action === "setup") { if(!actorIsLasgum_(actor)) throw new Error('AUTH_FORBIDDEN: Setup database hanya untuk LASGUM.'); return lasgumJson_(setupDatabase()); }
  if (action === "create") return lasgumJson_(apiCreate_(body.table, body.record, body.deviceId, actor));
  if (action === "update") return lasgumJson_(apiUpdate_(body.table, body.record, body.deviceId, actor));
  if (action === "delete") return lasgumJson_(apiDelete_(body.table, body.id, body.deviceId, actor));
  if (action === "sync") return lasgumJson_(apiSync_(body, actor));
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

function testSchemaV6() {
  const health = healthCheck();
  const meta = { ok: true, schemaVersion: LASGUM_SCHEMA_VERSION, tables: Object.keys(LASGUM_SHEETS) };
  if (health.schemaVersion !== LASGUM_SCHEMA_VERSION) throw new Error('Health/schemaVersion tidak konsisten.');
  if (meta.schemaVersion !== LASGUM_SCHEMA_VERSION || meta.tables.length !== 15) throw new Error('Meta/schema V6 tidak konsisten.');
  const mp = LASGUM_SHEETS.MATA_PELAJARAN;
  ['cakupan','scopeVersion'].forEach(function(field){
    if (mp.indexOf(field) === -1) throw new Error('MATA_PELAJARAN belum memiliki kolom '+field+'.');
  });
  Logger.log(JSON.stringify({ok:true,test:'SCHEMA_V6',schemaVersion:LASGUM_SCHEMA_VERSION,tableCount:meta.tables.length,message:'Schema V5 aktif dan 15 tabel terdaftar.'},null,2));
  return {ok:true,test:'SCHEMA_V6',schemaVersion:LASGUM_SCHEMA_VERSION,tableCount:meta.tables.length};
}

// Alias kompatibilitas untuk menu/test lama.
function testSchemaV3() { return testSchemaV6(); }


function testSiswaSchemaV6() {
  const sheet = lasgumGetSheet_("SISWA");
  const expected = LASGUM_SHEETS.SISWA;
  const actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0].map(String);

  const mismatch = expected.filter(function(header, index) {
    return actual[index] !== header;
  });

  if (mismatch.length) {
    throw new Error("Header SISWA belum sesuai V6. Jalankan setupDatabase().");
  }

  return { ok: true, test: "SISWA_SCHEMA_V6", headers: expected };
}

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

function testNormalizationDatabaseV6() {
  const deviceId = "TEST_V6_SIMPLE_STUDENT";
  const id = lasgumId_();
  const input = {
    id: id,
    nama: "SISWA TEST V6",
    nis: "9001",
    nisn: "1234567890",
    kelas: "3"
  };

  const created = apiCreate_("SISWA", input, deviceId);
  const read = apiGet_("SISWA", { id: id });

  if (!created.ok || !read.ok || read.count !== 1) {
    throw new Error("TEST SISWA V6 gagal.");
  }

  const row = read.data[0];
  if (row.nama !== input.nama ||
      String(row.nis) !== input.nis ||
      String(row.nisn) !== input.nisn ||
      String(row.kelas) !== input.kelas) {
    throw new Error("Data SISWA V6 tidak sesuai.");
  }

  apiDelete_("SISWA", id, deviceId);
  return { ok: true, test: "NORMALISASI_SISWA_V6", data: row };
}

function testIdempotentCreate() {
  const deviceId = "TEST_IDEMPOTENT_V6";
  const id = lasgumId_();
  const record = {
    id: id,
    nama: "SISWA TEST IDEMPOTENT V6",
    nis: "IDEMP001",
    nisn: "IDEMPNISN001",
    kelas: "2"
  };

  const first = apiCreate_("SISWA", record, deviceId);
  const second = apiCreate_("SISWA", record, deviceId);
  const read = apiGet_("SISWA", { id: id });

  if (!first.ok || !second.ok || !second.duplicate || read.count !== 1) {
    throw new Error("CREATE idempotent gagal.");
  }

  apiDelete_("SISWA", id, deviceId);
  return { ok: true, test: "IDEMPOTENT_CREATE_V6", count: read.count };
}

