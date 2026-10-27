// =========================================================================
// CẤU HÌNH HỆ THỐNG
// =========================================================================
var TARGET_SPREADSHEET_ID = "1uVuuKfxsTqpfKQ79WVzCZEeKcbCbCb2AEzB8Q4BYWtg";
var TIMEZONE = "Asia/Ho_Chi_Minh";

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss ? ss : SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------------------
// 1. XỬ LÝ GET & JSONP (VƯỢT QUA CORS 100% CHO VERCEL / GITHUB PAGES)
// -------------------------------------------------------------------------
function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAppConfig";
    var resData = { status: "error", message: "Action không hợp lệ" };

    if (action === "getAppConfig") {
      resData = { status: "success", data: getAppConfig() };
    } else if (action === "checkInEmployee") {
      var code = (e.parameter.empCode || "").toString().trim();
      var mType = (e.parameter.mealType || "MAIN").toString().trim();
      var checkRes = checkInEmployee(code, mType);
      resData = { status: "success", data: checkRes };
    } else if (action === "submitMainMeal") {
      var pMain = JSON.parse(e.parameter.payload || "{}");
      var mainRes = submitMainMeal(pMain);
      resData = { status: "success", data: mainRes };
    } else if (action === "submitExtraMealBatch") {
      var pExtra = JSON.parse(e.parameter.payload || "{}");
      var extraRes = submitExtraMealBatch(pExtra);
      resData = { status: "success", data: extraRes };
    }

    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(resData) + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } else {
      return ContentService.createTextOutput(JSON.stringify(resData))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    var errObj = { status: "error", message: err.toString() };
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(errObj) + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(errObj))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------------------
// 2. XỬ LÝ POST (DÀNH CHO ĐỒNG BỘ PYTHON & HỖ TRỢ DỰ PHÒNG)
// -------------------------------------------------------------------------
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "error", message: "postData rỗng" });
    }
    var body = JSON.parse(e.postData.contents);

    // Đồng bộ từ Python
    if (body.data && Array.isArray(body.data)) {
      return handlePythonSync(body.data);
    }

    var action = body.action;
    if (action === "submitMainMeal") {
      return responseJSON({ status: "success", data: submitMainMeal(body.payload) });
    }
    if (action === "submitExtraMealBatch") {
      return responseJSON({ status: "success", data: submitExtraMealBatch(body.payload) });
    }
    if (action === "checkInEmployee") {
      return responseJSON({ status: "success", data: checkInEmployee(body.empCode, body.mealType) });
    }
    return responseJSON({ status: "error", message: "Action không hợp lệ: " + action });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

// -------------------------------------------------------------------------
// CÁC HÀM NGHIỆP VỤ BẢNG TÍNH
// -------------------------------------------------------------------------
function getEmployeeMap() {
  var map = {};
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Danh_Sach_NV");
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < data.length; i++) {
        var code = (data[i][0] || "").toString().trim().toUpperCase();
        var name = (data[i][1] || "").toString().trim();
        var dept = (data[i][2] || "").toString().trim();
        if (code) map[code] = { name: name, dept: dept };
      }
    }
  } catch (e) {
    Logger.log("Lỗi nạp Danh_Sach_NV: " + e);
  }
  return map;
}

function getEmployeeName(code) {
  if (!code) return "";
  var map = getEmployeeMap();
  var emp = map[code.toString().trim().toUpperCase()];
  return emp ? emp.name : "";
}

function isValidEmpCode(code) {
  if (!code) return false;
  return /^[A-Z0-9]{3}-[0-9]{4}$/i.test(code.toString().trim());
}

function handlePythonSync(empList) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Danh_Sach_NV") || ss.insertSheet("Danh_Sach_NV");
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([["Mã Nhân Viên", "Họ Và Tên", "Tên Bộ Phận"]]);
  sheet.getRange("A1:C1").setBackground("#1F4E79").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(2, 1, empList.length, 3).setValues(empList);
  SpreadsheetApp.flush();
  return responseJSON({ status: "success", message: "Đã lưu Danh_Sach_NV thành công!", count: empList.length });
}

function getAppConfig() {
  var now = new Date();
  var hour = parseInt(Utilities.formatDate(now, TIMEZONE, "HH"), 10);
  var minute = parseInt(Utilities.formatDate(now, TIMEZONE, "mm"), 10);
  var currentMinutes = hour * 60 + minute;

  var employees = getEmployeeMap();
  var deptSet = {};
  Object.keys(employees).forEach(function(code) {
    var dept = employees[code].dept;
    if (dept) deptSet[dept] = true;
  });

  return {
    isMainLocked: currentMinutes >= (9 * 60 + 30),          // Khóa lúc 09:30
    isExtraLocked: currentMinutes >= (15 * 60 + 15),   // Khóa lúc 15:15
    currentTime: Utilities.formatDate(now, TIMEZONE, "HH:mm"),
    employees: employees,
    departments: Object.keys(deptSet).sort()
  };
}

function submitMainMeal(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var config = getAppConfig();
    if (config.isMainLocked) throw new Error("ĐÃ QUÁ 09:30! Cổng báo suất chính đã đóng.");
    if (!isValidEmpCode(data.empCode)) throw new Error("Mã nhân viên người báo không đúng định dạng!");

    var c1 = parseInt(data.ca1, 10) || 0;
    var c4 = parseInt(data.ca4, 10) || 0;
    var c2 = parseInt(data.ca2, 10) || 0;
    var c3 = parseInt(data.ca3, 10) || 0;
    var total = c1 + c4 + c2 + c3;
    if (total <= 0) throw new Error("Vui lòng nhập số suất cho ít nhất một ca.");

    var repName = (data.reporterName || "").trim() || getEmployeeName(data.empCode);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Dat_Suat_Chinh');
    var now = new Date();

    sheet.appendRow([now, data.department.trim(), data.empCode.toUpperCase().trim(), repName, c1, c4, c2, c3, total, (data.note || "").trim()]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// XỬ LÝ LƯU ĐĂNG KÝ TĂNG CA (CÓ CA ĂN, KHÔNG CHẶN TRÙNG, VẪN KHÓA LÚC 15:15)
// -------------------------------------------------------------------------
function submitExtraMealBatch(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var config = getAppConfig();
    // Vẫn chặn khóa cổng lúc 15:15
    if (config.isExtraLocked) throw new Error("ĐÃ QUÁ 15:15! Cổng đăng ký suất tăng ca đã đóng.");
    if (!isValidEmpCode(payload.reporterCode)) throw new Error("Mã NV người lập không đúng định dạng!");
    if (!payload.list || payload.list.length === 0) throw new Error("Danh sách nhân viên tăng ca đang trống.");

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Dat_Suat_Phu');
    if (!sheet) {
      sheet = ss.insertSheet('Dat_Suat_Phu');
    }

    // Nếu bảng chưa có tiêu đề hoặc là trang mới, tạo tiêu đề 9 cột
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Thời gian", "Bộ Phận", "Mã Người Lập", "Tên Người Lập", "Mã Nhân Viên", "Họ Và Tên", "Ca Ăn", "Món Ăn", "Ghi Chú"]);
      sheet.getRange("A1:I1").setBackground("#B26A00").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      SpreadsheetApp.flush();
    }

    var repName = (payload.reporterName || "").trim() || getEmployeeName(payload.reporterCode);
    var now = new Date();
    var rowsToAdd = [];

    for (var j = 0; j < payload.list.length; j++) {
      var p = payload.list[j];
      var code = (p.code || "").toUpperCase().trim();
      if (!isValidEmpCode(code)) throw new Error("Dòng " + (j + 1) + ": Mã NV '" + code + "' không đúng chuẩn!");
      if (!p.name || !p.name.trim()) throw new Error("Dòng " + (j + 1) + ": Chưa có họ tên nhân viên!");

      // Lưu 9 cột (có thêm cột Ca Ăn - p.mealShift)
      rowsToAdd.push([
        now, 
        payload.department.trim(), 
        payload.reporterCode.toUpperCase().trim(), 
        repName, 
        code, 
        p.name.trim(), 
        p.mealShift || "Ca 2 (17:00 - 18:30)", 
        p.dish, 
        (p.note || "").trim()
      ]);
    }

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, 9).setValues(rowsToAdd);
    sheet.getRange(startRow, 1, rowsToAdd.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    SpreadsheetApp.flush();

    return { success: true, count: rowsToAdd.length };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// ĐỐI SOÁT ĐIỂM DANH ĂN (CẬP NHẬT ĐỌC THEO CẤU TRÚC 9 CỘT CỦA BẢNG TĂNG CA)
// -------------------------------------------------------------------------
function checkInEmployee(empCode, mealType) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!empCode) throw new Error("Mã nhân viên không được để trống!");
    mealType = mealType || "MAIN";

    var code = empCode.toString().trim().toUpperCase();
    var empMap = getEmployeeMap();
    var emp = empMap[code];

    if (!emp) {
      return { success: false, status: "NOT_FOUND", message: "Mã NV [" + code + "] không tồn tại trong danh bạ!" };
    }

    var ss = getSpreadsheet();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy");
    var timeStr = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");
    var mealLabel = (mealType === "EXTRA") ? "Suất Tăng Ca" : "Suất Chính";
    var dishDetail = (mealType === "EXTRA") ? "" : "Cơm ca";

    // Đối soát danh sách tăng ca
    if (mealType === "EXTRA") {
      var sExtra = ss.getSheetByName('Dat_Suat_Phu');
      var isRegistered = false;
      if (sExtra && sExtra.getLastRow() > 1) {
        var extraRecords = sExtra.getRange(2, 1, sExtra.getLastRow() - 1, 8).getValues();
        for (var i = 0; i < extraRecords.length; i++) {
          var rowDate = extraRecords[i][0];
          var dStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, TIMEZONE, "dd/MM/yyyy") : (rowDate || "").toString().substring(0, 10);
          var recCode = (extraRecords[i][4] || "").toString().trim().toUpperCase();
          if (dStr === todayStr && recCode === code) {
            isRegistered = true;
            var caAn = (extraRecords[i][6] || "").toString().trim();
            var monAn = (extraRecords[i][7] || "Suất tăng ca").toString().trim();
            dishDetail = caAn ? (caAn + " - " + monAn) : monAn;
            break;
          }
        }
      }

      if (!isRegistered) {
        return { success: false, status: "NOT_REGISTERED", empCode: code, name: emp.name, dept: emp.dept, message: "KHÔNG CÓ TÊN TRONG DANH SÁCH TĂNG CA HÔM NAY!" };
      }
    }

    var sheet = ss.getSheetByName("Diem_Danh_An");
    if (!sheet) {
      sheet = ss.insertSheet("Diem_Danh_An");
      sheet.getRange(1, 1, 1, 7).setValues([["Thời gian quét", "Ngày", "Mã Nhân Viên", "Họ Và Tên", "Bộ Phận", "Bữa Ăn", "Món / Chi Tiết"]]);
      sheet.getRange("A1:G1").setBackground("#2E7D32").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      SpreadsheetApp.flush();
    }

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var checkData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var k = 0; k < checkData.length; k++) {
        var rowD = checkData[k][1];
        var cDateStr = (rowD instanceof Date) ? Utilities.formatDate(rowD, TIMEZONE, "dd/MM/yyyy") : (rowD || "").toString();
        var cCode = (checkData[k][2] || "").toString().trim().toUpperCase();
        var cMeal = (checkData[k][5] || "").toString().trim();

        if (cDateStr === todayStr && cCode === code && cMeal === mealLabel) {
          var prevTime = (checkData[k][0] instanceof Date) ? Utilities.formatDate(checkData[k][0], TIMEZONE, "HH:mm:ss") : "";
          return { success: false, status: "DUPLICATE", empCode: code, name: emp.name, dept: emp.dept, mealLabel: mealLabel, message: "ĐÃ ĂN " + mealLabel.toUpperCase() + " HÔM NAY RỒI!" + (prevTime ? " (Lúc " + prevTime + ")" : "") };
        }
      }
    }

    sheet.appendRow([now, todayStr, code, emp.name, emp.dept, mealLabel, dishDetail]);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(newRow, 2).setNumberFormat("@");
    SpreadsheetApp.flush();

    return { success: true, status: "OK", empCode: code, name: emp.name, dept: emp.dept, mealLabel: mealLabel, dishDetail: dishDetail, time: timeStr };
  } finally {
    lock.releaseLock();
  }
}