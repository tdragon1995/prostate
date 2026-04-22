// ============================================================
// PROSTATE SCREENING DATABASE — Apps Script Web App
// Sheet ID: 1330j8eh2Qqce3zQrxUtMINrkHVp3DQc6N0_LY5aWozU
// Deploy: Extensions → Apps Script → Deploy → Web App
//         Execute as: Me | Who has access: Anyone
// ============================================================

const SHEET_ID   = "1330j8eh2Qqce3zQrxUtMINrkHVp3DQc6N0_LY5aWozU";
const SHEET_NAME = "Raw Data";

const HEADERS = [
  "STT",
  "Mã Lượt Khám",
  "Họ và Tên",
  "Ngày Sinh",
  "Tuổi",
  "Dân Tộc",
  "Địa chỉ thường trú",
  "Địa chỉ tạm trú",
  "Trình độ học vấn",
  "Nghề nghiệp",
  "Tiền sử gia đình UTTTL",
  "Tiền sử gia đình ung thư khác",
  "Tiền sử phì đại tuyến tiền liệt",
  "Tiền sử viêm tuyến tiền liệt",
  "Tiền sử nhiễm trùng tiểu",
  "Đái tháo đường",
  "Tăng huyết áp",
  "Rối loạn chuyển hóa",
  "Tiền sử dùng thuốc",
  "Tiểu đêm",
  "Dòng tiểu yếu",
  "Tiểu khó",
  "Tiểu nhỏ giọt",
  "Tiểu gấp",
  "Tiểu không hết",
  "Rối loạn cương",
  "Đau vùng chậu/sinh dục",
  "Điểm IPSS",
  "Nhận thức trước tầm soát",
  "Lý do tham gia",
  "Mức độ e ngại",
  "Mức độ thoải mái",
  "Thể tích tuyến tiền liệt",
  "Mật độ PSA",
  "Kết quả thăm trực tràng",
  "Có chuyển tuyến không",
  "Chuyên khoa chuyển",
  "Có sinh thiết không",
  "Kết quả sinh thiết",
  "Điểm Gleason",
  "Số mẫu dương tính",
  "Kết quả hình ảnh",
  "Nhân sự thực hiện",
  "Ngày nhập liệu",
];

// ── Shared: get or create sheet ──────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// ── Shared: write + format header row ───────────────────────
function writeHeaders(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
  hdr.setBackground("#1a73e8");
  hdr.setFontColor("#ffffff");
  hdr.setFontWeight("bold");
  hdr.setWrap(true);
  hdr.setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 48);
  sheet.autoResizeColumns(1, HEADERS.length);
}

// ── RUN THIS ONCE to fix mismatched headers ──────────────────
// Select resetSheet in the editor dropdown → click Run
function resetSheet() {
  const sheet = getSheet();

  // Delete only row 1 (header) — data rows are untouched
  if (sheet.getLastRow() >= 1) {
    sheet.deleteRow(1);
  }

  // Insert a fresh header row at the top
  sheet.insertRowBefore(1);
  writeHeaders(sheet);

  Logger.log("✅ Headers reset. " + sheet.getLastRow() + " rows total (including new header).");
}

// ── GET: health check & fetch data ──────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'health';

  if (action === 'fetchData') {
    try {
      const sheet = getSheet();
      const data = sheet.getDataRange().getValues();

      if (data.length <= 1) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: "ok", rows: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const headers = data[0];
      const rows = data.slice(1).map(r => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = r[i] || '');
        return obj;
      });

      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default health check
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Prostate Screening API hoạt động" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST: receive data from form (with update support) ───────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    // Write headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      writeHeaders(sheet);
    }

    // Timestamp (Vietnam time)
    const now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");

    // Check if patient already exists (by STT if provided, or skip if new)
    let stt = data["STT"];
    let existingRowIndex = null;

    if (stt) {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === stt) {
          existingRowIndex = i + 1; // Google Sheets uses 1-based indexing
          break;
        }
      }
    }

    // Build row array (matching HEADERS order)
    const row = [
      stt || sheet.getLastRow(),
      data["Mã Lượt Khám"]                      || "",
      data["Họ và Tên"]                         || "",
      data["Ngày Sinh"]                         || "",
      data["Tuổi"]                              || "",
      data["Dân Tộc"]                           || "",
      data["Địa chỉ thường trú"]                || "",
      data["Địa chỉ tạm trú"]                   || "",
      data["Trình độ học vấn"]                  || "",
      data["Nghề nghiệp"]                       || "",
      data["Tiền sử gia đình UTTTL"]            || "",
      data["Tiền sử gia đình ung thư khác"]     || "",
      data["Tiền sử phì đại tuyến tiền liệt"]   || "",
      data["Tiền sử viêm tuyến tiền liệt"]      || "",
      data["Tiền sử nhiễm trùng tiểu"]          || "",
      data["Đái tháo đường"]                    || "",
      data["Tăng huyết áp"]                     || "",
      data["Rối loạn chuyển hóa"]               || "",
      data["Tiền sử dùng thuốc"]                || "",
      data["Tiểu đêm"]                          || "",
      data["Dòng tiểu yếu"]                     || "",
      data["Tiểu khó"]                          || "",
      data["Tiểu nhỏ giọt"]                     || "",
      data["Tiểu gấp"]                          || "",
      data["Tiểu không hết"]                    || "",
      data["Rối loạn cương"]                    || "",
      data["Đau vùng chậu/sinh dục"]            || "",
      data["Điểm IPSS"]                         || "",
      data["Nhận thức trước tầm soát"]          || "",
      data["Lý do tham gia"]                    || "",
      data["Mức độ e ngại"]                     || "",
      data["Mức độ thoải mái"]                  || "",
      data["Thể tích tuyến tiền liệt"]          || "",
      data["Mật độ PSA"]                        || "",
      data["Kết quả thăm trực tràng"]           || "",
      data["Có chuyển tuyến không"]             || "",
      data["Chuyên khoa chuyển"]                || "",
      data["Có sinh thiết không"]               || "",
      data["Kết quả sinh thiết"]                || "",
      data["Điểm Gleason"]                      || "",
      data["Số mẫu dương tính"]                 || "",
      data["Kết quả hình ảnh"]                  || "",
      data["Nhân sự thực hiện"]                 || "",
      now,
    ];

    // Update existing row or append new
    if (existingRowIndex) {
      sheet.getRange(existingRowIndex, 1, 1, row.length).setValues([row]);
      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", stt: stt, timestamp: now, action: "updated" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      sheet.appendRow(row);
      const newStt = sheet.getLastRow() - 1; // Subtract 1 for header
      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", stt: newStt, timestamp: now, action: "created" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
