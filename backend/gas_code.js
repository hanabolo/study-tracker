// ============================================================
// 設定：デプロイ後にこの値を自分のAPIキーに変更してください
// ============================================================
const API_KEY = "YOUR_SECRET_API_KEY_HERE"; // 例: "family2024secret"

// スプレッドシートのシート名
const SHEET_NAME = "records";

// ============================================================
// GETリクエスト処理（データ取得・バックアップ）
// ============================================================
function doGet(e) {
  try {
    const key = e.parameter.apiKey;
    if (key !== API_KEY) {
      return jsonResponse({ error: "Unauthorized" }, 401);
    }

    const action = e.parameter.action;
    const sheet = getSheet();

    if (action === "getRecords") {
      const records = getAllRecords(sheet);
      return jsonResponse({ records });
    }

    if (action === "getSummary") {
      const records = getAllRecords(sheet);
      const summary = calcSummary(records);
      return jsonResponse(summary);
    }

    if (action === "backup") {
      const records = getAllRecords(sheet);
      return jsonResponse({ records, exportedAt: new Date().toISOString() });
    }

    return jsonResponse({ error: "Unknown action" }, 400);

  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ============================================================
// POSTリクエスト処理（データ書き込み・削除・リストア）
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const key = body.apiKey;
    if (key !== API_KEY) {
      return jsonResponse({ error: "Unauthorized" }, 401);
    }

    const action = body.action;
    const sheet = getSheet();

    // 学習時間 or 余暇利用時間を記録
    if (action === "addRecord") {
      const { date, type, minutes, note } = body;
      // type: "study" or "leisure_used"
      const id = Utilities.getUuid();
      const now = new Date().toISOString();
      sheet.appendRow([id, date, type, minutes, note || "", now]);
      return jsonResponse({ success: true, id });
    }

    // レコード削除
    if (action === "deleteRecord") {
      const { id } = body;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ success: true });
        }
      }
      return jsonResponse({ error: "Record not found" }, 404);
    }

    // バックアップからリストア
    if (action === "restore") {
      const { records } = body;
      // 既存データを全削除（ヘッダー以外）
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      // 全レコードを挿入
      records.forEach(r => {
        sheet.appendRow([r.id, r.date, r.type, r.minutes, r.note || "", r.createdAt]);
      });
      return jsonResponse({ success: true, restored: records.length });
    }

    return jsonResponse({ error: "Unknown action" }, 400);

  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ============================================================
// ヘルパー関数
// ============================================================

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["id", "date", "type", "minutes", "note", "createdAt"]);
  }
  return sheet;
}

function getAllRecords(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj.minutes = Number(obj.minutes);
    return obj;
  });
}

function calcSummary(records) {
  // 総学習時間（分）
  const totalStudy = records
    .filter(r => r.type === "study")
    .reduce((sum, r) => sum + r.minutes, 0);

  // 余暇可能時間 = 学習時間 × 1/4
  const totalLeisureAvailable = Math.floor(totalStudy / 4);

  // 余暇利用時間
  const totalLeisureUsed = records
    .filter(r => r.type === "leisure_used")
    .reduce((sum, r) => sum + r.minutes, 0);

  // 余暇残高
  const leisureBalance = totalLeisureAvailable - totalLeisureUsed;

  return {
    totalStudy,
    totalLeisureAvailable,
    totalLeisureUsed,
    leisureBalance,
  };
}

function jsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
