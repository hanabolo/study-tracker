// ============================================================
// 設定：デプロイ後にこの値を自分のAPIキーに変更してください
// ============================================================
const API_KEY = "YOUR_SECRET_API_KEY_HERE"; // 例: "family2024secret"

// スプレッドシートのシート名
const SHEET_NAME = "records";

// 列定義（順序はシートの並びと一致させること）
const HEADERS = ["id", "date", "type", "minutes", "note", "createdAt"];
const NUM_COLS = HEADERS.length;
const TZ = "Asia/Tokyo";

// ============================================================
// GETリクエスト処理
// ============================================================
function doGet(e) {
  try {
    if (e.parameter.apiKey !== API_KEY) {
      return jsonResponse({ error: "Unauthorized" });
    }
    return handleAction(e.parameter.action, e.parameter);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// POSTリクエスト処理
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.apiKey !== API_KEY) {
      return jsonResponse({ error: "Unauthorized" });
    }
    return handleAction(body.action, body);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// アクション本体（GET/POSTで共通）
// ============================================================
function handleAction(action, p) {
  const sheet = getSheet();

  if (action === "getRecords") {
    return jsonResponse({ records: getAllRecords(sheet) });
  }

  if (action === "getSummary") {
    return jsonResponse(calcSummary(getAllRecords(sheet)));
  }

  if (action === "backup") {
    return jsonResponse({
      records: getAllRecords(sheet),
      exportedAt: new Date().toISOString(),
    });
  }

  // 追加した1件をそのまま返す。クライアントは再取得なしで画面を更新できる
  if (action === "addRecord") {
    return withLock(() => {
      const record = {
        id: Utilities.getUuid(),
        date: p.date,
        type: p.type,
        minutes: Number(p.minutes),
        note: p.note || "",
        createdAt: new Date().toISOString(),
      };
      sheet.appendRow(HEADERS.map(h => record[h]));
      return jsonResponse({ success: true, id: record.id, record });
    });
  }

  if (action === "deleteRecord") {
    return withLock(() => {
      const deleted = deleteRecordById(sheet, p.id);
      return deleted
        ? jsonResponse({ success: true, id: p.id })
        : jsonResponse({ error: "Record not found" });
    });
  }

  if (action === "restore") {
    return withLock(() => {
      const records = typeof p.records === "string" ? JSON.parse(p.records) : p.records;
      const restored = restoreRecords(sheet, records || []);
      return jsonResponse({ success: true, restored });
    });
  }

  return jsonResponse({ error: "Unknown action" });
}

// ============================================================
// シート操作
// ============================================================

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function getAllRecords(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // getDataRange()ではなく必要な範囲だけ読む（余分な列・書式のみの行を拾わない）
  const rows = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  const out = new Array(rows.length);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    out[i] = {
      id: row[0],
      date: cellToString(row[1]),
      type: row[2],
      minutes: Number(row[3]),
      note: row[4],
      createdAt: cellToString(row[5]),
    };
  }
  return out;
}

// 削除対象を探すのにid列（A列）だけ読む。全列読み込みに比べて転送量が1/6になる
function deleteRecordById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(i + 2); // ヘッダー行ぶん+1、0始まりぶん+1
      return true;
    }
  }
  return false;
}

// appendRowをループで呼ぶと1件ごとにシートへ往復するため、setValuesで一括書き込みする
function restoreRecords(sheet, records) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  if (!records.length) return 0;

  const values = records.map(r => [
    r.id, r.date, r.type, Number(r.minutes), r.note || "", r.createdAt,
  ]);
  sheet.getRange(2, 1, values.length, NUM_COLS).setValues(values);
  return values.length;
}

// ============================================================
// ヘルパー関数
// ============================================================

// 同時書き込みで行がずれたり二重登録されたりするのを防ぐ
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function cellToString(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : v;
}

function calcSummary(records) {
  let totalStudy = 0;
  let totalLeisureUsed = 0;
  for (let i = 0; i < records.length; i++) {
    const r = records[i];
    if (r.type === "study") totalStudy += r.minutes;
    else if (r.type === "leisure_used") totalLeisureUsed += r.minutes;
  }

  // 余暇可能時間 = 学習時間 × 1/4
  const totalLeisureAvailable = Math.floor(totalStudy / 4);

  return {
    totalStudy,
    totalLeisureAvailable,
    totalLeisureUsed,
    leisureBalance: totalLeisureAvailable - totalLeisureUsed,
  };
}

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
