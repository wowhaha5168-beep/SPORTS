// ======================================================
// 員工運動報名系統 - Google Apps Script v2
// 格式：每個運動獨立工作表，場次橫向排列
// ======================================================

const SHEET_ID = "1VqSECXOB15jpura28xAmtBSQefQO5OX60pVJSrsIIaU";

// 個別月份名額調整（覆蓋預設名額）：{ 運動: { "年-月": 名額 } }
const CAPACITY_OVERRIDES = {
  tennis: { "2026-10": 8 },
};
function getCapacity(sport, year, month, defaultCap) {
  const o = CAPACITY_OVERRIDES[sport];
  const v = o && o[`${year}-${month}`];
  return v != null ? v : defaultCap;
}

const SPORT_CONFIG = {
  yoga: {
    label: "瑜伽", sheetName: "瑜伽",
    sessions: (year, month) => {
      const days = []; const names = ["日","一","二","三","四","五","六"];
      const dim = new Date(year, month, 0).getDate();
      for (let d = 1; d <= dim; d++) {
        const day = new Date(year, month - 1, d).getDay();
        const dd = `${month}/${String(d).padStart(2,"0")}`;
        if (day === 3) days.push({ id: `yoga-${d}-1730`, label: `${dd}（${names[day]}）17:30～18:30`, capacity: 20 });
      }
      return days;
    }
  },
  badminton: {
    label: "羽球", sheetName: "羽球",
    sessions: (year, month) => {
      const days = []; const names = ["日","一","二","三","四","五","六"];
      const dim = new Date(year, month, 0).getDate();
      for (let d = 1; d <= dim; d++) {
        const day = new Date(year, month - 1, d).getDay();
        const dd = `${month}/${String(d).padStart(2,"0")}`;
        if (day === 2) {
          days.push({ id: `badminton-${d}-0900`, label: `${dd}（${names[day]}）09:00～11:00`, capacity: 8 });
          days.push({ id: `badminton-${d}-1700`, label: `${dd}（${names[day]}）17:00～19:00`, capacity: 20 });
        }
      }
      return days;
    }
  },
  tennis: {
    label: "網球", sheetName: "網球",
    sessions: (year, month) => {
      const days = []; const names = ["日","一","二","三","四","五","六"];
      const dim = new Date(year, month, 0).getDate();
      for (let d = 1; d <= dim; d++) {
        const day = new Date(year, month - 1, d).getDay();
        const dd = `${month}/${String(d).padStart(2,"0")}`;
        if (day === 3) days.push({ id: `tennis-${d}-1900`, label: `${dd}（${names[day]}）19:00～21:00`, capacity: getCapacity("tennis", year, month, 10) });
      }
      return days;
    }
  }
};

const SPORTS = ["yoga", "badminton", "tennis"];
// 「剩餘名額」分頁的欄位順序（由左至右）
const REMAIN_ORDER = ["badminton", "yoga", "tennis"];
const REMAIN_SHEET_SUFFIX = "剩餘名額";

// ── 找出指定年月專屬的報名表檔案，不存在則回傳 null（不會新建）──
function findMonthlySpreadsheet(year, month) {
  const fileName = `${year}年${month}月運動報名表`;
  const baseFile = DriveApp.getFileById(SHEET_ID);
  const parents = baseFile.getParents();
  const folder = parents.hasNext() ? parents.next() : null;

  const files = folder ? folder.getFilesByName(fileName) : DriveApp.getFilesByName(fileName);
  return files.hasNext() ? SpreadsheetApp.open(files.next()) : null;
}

// ── 取得（或新建）指定年月專屬的報名表檔案 ──
// 例如：2026年10月運動報名表，與「2026年9月運動報名表」放在同一個資料夾
function getMonthlySpreadsheet(year, month) {
  const existing = findMonthlySpreadsheet(year, month);
  if (existing) return existing;

  const fileName = `${year}年${month}月運動報名表`;
  const baseFile = DriveApp.getFileById(SHEET_ID);
  const parents = baseFile.getParents();
  const folder = parents.hasNext() ? parents.next() : null;

  const newSs = SpreadsheetApp.create(fileName);
  const newFile = DriveApp.getFileById(newSs.getId());
  if (folder) {
    folder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
  }
  return newSs;
}

// ── 新建檔案時會多一個預設的空白分頁，等其他分頁建好後清掉 ──
function removeDefaultSheet(ss) {
  if (ss.getSheets().length <= 1) return;
  ["工作表1", "Sheet1"].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
}

function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else {
      return res({ error: "No data received" });
    }
    return handleAction(data);
  } catch (err) {
    return res({ error: err.message });
  }
}

function doGet(e) {
  try {
    let result = { status: "ok" };
    if (e.parameter && e.parameter.data) {
      const data = JSON.parse(e.parameter.data);
      const output = handleAction(data);
      const json = output.getContent();
      result = JSON.parse(json);
    }
    // JSONP support
    if (e.parameter && e.parameter.callback) {
      return ContentService
        .createTextOutput(`${e.parameter.callback}(${JSON.stringify(result)})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return res(result);
  } catch (err) {
    return res({ error: err.message });
  }
}

function handleAction(data) {
  const action = data.action;
  if (action === "addReg")    return addReg(data);
  if (action === "updateReg") return updateReg(data);
  if (action === "deleteReg") return deleteReg(data);
  if (action === "getRegs")   return getRegs(data);
  if (action === "getPassword")  return getPassword();
  if (action === "savePassword") return savePassword(data.password);
  if (action === "rebuildRemaining") return rebuildRemaining(data);
  return res({ error: "Unknown action" });
}

// ── 手動補建/刷新指定月份的「剩餘名額」分頁（用於報名已截止、不會再觸發 addReg 的情況）──
function rebuildRemaining(data) {
  const year = parseInt(data.year, 10);
  const month = parseInt(data.month, 10);
  if (!year || !month) return res({ error: "缺少 year 或 month" });
  const ss = findMonthlySpreadsheet(year, month);
  if (!ss) return res({ error: "找不到該月份的報名表檔案" });
  buildRemainingSheet(ss, year, month);
  return res({ success: true, year, month });
}

// ── 取得密碼 ──
function getPassword() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("設定");
  if (!sheet) {
    sheet = ss.insertSheet("設定");
    sheet.getRange(1, 1).setValue("admin_password");
    sheet.getRange(1, 2).setValue("admin123");
  }
  const pw = sheet.getRange(1, 2).getValue();
  return res({ password: pw || "admin123" });
}

// ── 儲存密碼 ──
function savePassword(password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("設定");
  if (!sheet) {
    sheet = ss.insertSheet("設定");
    sheet.getRange(1, 1).setValue("admin_password");
  }
  sheet.getRange(1, 2).setValue(password);
  return res({ success: true });
}

// ── 取得報名資料（從總表）──
function getRegs(data) {
  const ss = findMonthlySpreadsheet(data.year, data.month);
  if (!ss) return res({ regs: [] });
  const sheetName = `${data.year}年${data.month}月總表`;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return res({ regs: [] });

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return res({ regs: [] });

  const headers = rows[0];
  const regs = rows.slice(1).filter(r => r[0] !== "").map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return res({ regs });
}

// ── 新增報名 ──
function addReg(data) {
  const year = data.year;
  const month = data.month;
  const ss = getMonthlySpreadsheet(year, month);

  // 確保所有工作表存在
  ensureAllSheets(ss, year, month);

  // 寫入總表
  const summarySheet = ss.getSheetByName(`${year}年${month}月總表`);
  const rows = summarySheet.getDataRange().getValues();
  const nameCol = rows[0].indexOf("姓名");

  // 檢查重複
  if (rows.slice(1).some(r => r[nameCol] === data.name)) {
    return res({ error: "此姓名本月已報名" });
  }

  const id = new Date().getTime().toString();
  const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm");

  summarySheet.appendRow([
    id, data.name, data.phone, data.department,
    (data.yoga || []).join("、"),
    (data.badminton || []).join("、"),
    (data.tennis || []).join("、"),
    now
  ]);

  // 寫入各運動工作表
  SPORTS.forEach(sport => {
    const cfg = SPORT_CONFIG[sport];
    const sportSheet = ss.getSheetByName(`${year}年${month}月${cfg.sheetName}`);
    const selected = data[sport] || [];
    const noneId = `${sport}-none`;
    if (selected.includes(noneId) || selected.length === 0) return;

    const sessions = cfg.sessions(year, month);
    // 找到各場次的欄位位置
    const headerRow1 = sportSheet.getRange(1, 1, 1, sportSheet.getLastColumn()).getValues()[0];

    selected.forEach(sessionId => {
      const session = sessions.find(s => s.id === sessionId);
      if (!session) return;

      // 找到這個場次的起始欄
      let startCol = -1;
      for (let c = 0; c < headerRow1.length; c++) {
        if (headerRow1[c] === session.label) { startCol = c + 1; break; }
      }
      if (startCol === -1) return;

      // 找空行填入
      const colData = sportSheet.getRange(3, startCol, 50, 1).getValues();
      let emptyRow = -1;
      for (let r = 0; r < colData.length; r++) {
        if (colData[r][0] === "") { emptyRow = r + 3; break; }
      }
      if (emptyRow === -1) return;

      sportSheet.getRange(emptyRow, startCol).setValue(data.name);
      sportSheet.getRange(emptyRow, startCol + 1).setValue(data.phone);
      sportSheet.getRange(emptyRow, startCol + 2).setValue(data.department);
    });
  });

  // 更新「剩餘名額」分頁
  buildRemainingSheet(ss, year, month);

  return res({ success: true, id });
}

// ── 更新報名 ──
function updateReg(data) {
  // 先刪除舊資料再新增
  deleteReg({ id: data.id, year: data.year, month: data.month });
  data.action = "addReg";
  return addReg(data);
}

// ── 刪除報名 ──
function deleteReg(data) {
  const year = data.year;
  const month = data.month;
  const ss = findMonthlySpreadsheet(year, month);
  if (!ss) return res({ error: "找不到該月份的報名表檔案" });

  // 從總表找姓名
  const summarySheet = ss.getSheetByName(`${year}年${month}月總表`);
  if (!summarySheet) return res({ error: "找不到工作表" });

  const rows = summarySheet.getDataRange().getValues();
  const idCol = rows[0].indexOf("ID");
  const nameCol = rows[0].indexOf("姓名");
  let targetName = "";

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] == data.id) {
      targetName = rows[i][nameCol];
      summarySheet.deleteRow(i + 1);
      break;
    }
  }

  if (!targetName) return res({ error: "找不到此筆資料" });

  // 從各運動工作表刪除
  SPORTS.forEach(sport => {
    const cfg = SPORT_CONFIG[sport];
    const sportSheet = ss.getSheetByName(`${year}年${month}月${cfg.sheetName}`);
    if (!sportSheet) return;

    const lastCol = sportSheet.getLastColumn();
    const lastRow = sportSheet.getLastRow();
    if (lastRow < 3) return;

    const allData = sportSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    for (let r = 0; r < allData.length; r++) {
      for (let c = 0; c < allData[r].length; c++) {
        if (allData[r][c] === targetName) {
          sportSheet.getRange(r + 3, c + 1).setValue("");
          sportSheet.getRange(r + 3, c + 2).setValue("");
          sportSheet.getRange(r + 3, c + 3).setValue("");
        }
      }
    }
  });

  // 更新「剩餘名額」分頁
  buildRemainingSheet(ss, year, month);

  return res({ success: true });
}

// ── 建立所有工作表 ──
function ensureAllSheets(ss, year, month) {
  // 總表
  const summaryName = `${year}年${month}月總表`;
  if (!ss.getSheetByName(summaryName)) {
    const s = ss.insertSheet(summaryName);
    s.appendRow(["ID", "姓名", "手機", "部門", "瑜伽", "羽球", "網球", "報名時間"]);
    formatHeader(s, 1, 8, "#1D4ED8");
    s.setFrozenRows(1);
  }

  // 各運動工作表
  SPORTS.forEach(sport => {
    const cfg = SPORT_CONFIG[sport];
    const sheetName = `${year}年${month}月${cfg.sheetName}`;
    if (!ss.getSheetByName(sheetName)) {
      const s = ss.insertSheet(sheetName);
      buildSportSheet(s, cfg, year, month, sport);
    }
  });

  // 第五個分頁：剩餘名額（尚未有人報名時，先以滿額建立）
  if (!ss.getSheetByName(`${year}年${month}月${REMAIN_SHEET_SUFFIX}`)) {
    buildRemainingSheet(ss, year, month);
  }

  // 新檔案會多一個預設空白分頁，清掉它
  removeDefaultSheet(ss);
}

// ── 統計各場次目前已報名人數 ──
function getSessionCounts(ss, year, month) {
  const counts = {};
  const summarySheet = ss.getSheetByName(`${year}年${month}月總表`);
  if (!summarySheet) return counts;
  const rows = summarySheet.getDataRange().getValues();
  if (rows.length <= 1) return counts;

  const headers = rows[0];
  const colIdx = { yoga: headers.indexOf("瑜伽"), badminton: headers.indexOf("羽球"), tennis: headers.indexOf("網球") };
  rows.slice(1).forEach(row => {
    SPORTS.forEach(sport => {
      const idx = colIdx[sport];
      if (idx === -1) return;
      const val = row[idx];
      if (!val) return;
      String(val).split("、").forEach(id => {
        id = id.trim();
        if (!id || id.indexOf("-none") !== -1) return;
        counts[id] = (counts[id] || 0) + 1;
      });
    });
  });
  return counts;
}

// 依日期算出屬於當月第幾個「日～六」週（用來把同一週的場次排在同一列區塊）
function getWeekIndex(year, month, day) {
  const firstDow = new Date(year, month - 1, 1).getDay(); // 0=日
  return Math.floor((day - 1 + firstDow) / 7);
}

// ── 建立/更新「剩餘名額」分頁 ──
function buildRemainingSheet(ss, year, month) {
  const sheetName = `${year}年${month}月${REMAIN_SHEET_SUFFIX}`;
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(sheetName); }

  const counts = getSessionCounts(ss, year, month);

  // 依運動別，將場次依「週」分組
  const perSportWeeks = {};
  let maxWeek = 0;
  REMAIN_ORDER.forEach(sport => {
    const cfg = SPORT_CONFIG[sport];
    const sessions = cfg.sessions(year, month);
    const weeks = {};
    sessions.forEach(s => {
      const day = parseInt(s.id.split("-")[1], 10);
      const wk = getWeekIndex(year, month, day);
      if (!weeks[wk]) weeks[wk] = [];
      weeks[wk].push(s);
      if (wk > maxWeek) maxWeek = wk;
    });
    perSportWeeks[sport] = weeks;
  });

  // 表頭（第1列：項目／第2列：日期＋運動名稱）
  REMAIN_ORDER.forEach((sport, i) => {
    const dateCol = i * 2 + 1;   // A, C, E
    const itemCol = i * 2 + 2;  // B, D, F
    const cfg = SPORT_CONFIG[sport];
    sheet.getRange(1, itemCol).setValue("項目")
      .setBackground("#FFFF00").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(2, dateCol).setValue("日期")
      .setBackground("#00FFFF").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(2, itemCol).setValue(cfg.label)
      .setBackground("#FFF8DC").setFontWeight("bold").setHorizontalAlignment("center");
  });

  // 資料列：依週分組，每週區塊的列數 = 該週三項運動場次數的最大值
  let row = 3;
  for (let wk = 0; wk <= maxWeek; wk++) {
    let blockRows = 0;
    REMAIN_ORDER.forEach(sport => {
      const list = perSportWeeks[sport][wk] || [];
      if (list.length > blockRows) blockRows = list.length;
    });
    if (blockRows === 0) continue;

    REMAIN_ORDER.forEach((sport, i) => {
      const dateCol = i * 2 + 1;
      const list = perSportWeeks[sport][wk] || [];
      list.forEach((s, idx) => {
        const remain = Math.max(0, s.capacity - (counts[s.id] || 0));
        sheet.getRange(row + idx, dateCol).setValue(`${s.label} (尚餘: ${remain})`)
          .setFontColor("#0000FF").setBackground("#FDF6E3");
      });
    });

    // 每週區塊下方畫粗紅線分隔
    sheet.getRange(row + blockRows - 1, 1, 1, 6)
      .setBorder(false, false, true, false, false, false, "#CC0000", SpreadsheetApp.BorderStyle.SOLID_THICK);
    row += blockRows;
  }

  sheet.setColumnWidths(1, 6, 190);
  sheet.setFrozenRows(2);
}

function buildSportSheet(sheet, cfg, year, month, sport) {
  const sessions = cfg.sessions(year, month);
  if (sessions.length === 0) return;

  const colors = { yoga: "#7C3AED", badminton: "#059669", tennis: "#D97706" };
  const color = colors[sport] || "#1D4ED8";

  // 第一行：場次標題（每場佔3欄）
  // 第二行：匿稱、行動電話、部門
  let col = 1;
  sessions.forEach(session => {
    // 合併第一行的3欄作為場次標題
    sheet.getRange(1, col, 1, 3).merge();
    sheet.getRange(1, col).setValue(session.label);
    sheet.getRange(1, col, 1, 3).setBackground(color).setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");

    // 第二行子標題
    sheet.getRange(2, col).setValue("匿稱");
    sheet.getRange(2, col + 1).setValue("行動電話");
    sheet.getRange(2, col + 2).setValue("部門");
    sheet.getRange(2, col, 1, 3).setBackground("#E8F0FE").setFontWeight("bold")
      .setHorizontalAlignment("center");

    // 設定欄寬
    sheet.setColumnWidth(col, 80);
    sheet.setColumnWidth(col + 1, 110);
    sheet.setColumnWidth(col + 2, 70);

    // 淡色背景區域
    sheet.getRange(3, col, 30, 3).setBackground("#F8FAFC");

    col += 3;
  });

  sheet.setFrozenRows(2);
}

function formatHeader(sheet, row, cols, color) {
  const range = sheet.getRange(row, 1, 1, cols);
  range.setBackground(color).setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}

// ── 一次性搬遷用：把「運動」檔案裡指定年月的舊分頁搬到專屬檔案 ──
// 用法：Apps Script 編輯器上方函式下拉選單選 migrateOct2026，按執行（第一次會要求授權 Drive 權限）
function migrateMonthToOwnFile(year, month) {
  const oldSs = SpreadsheetApp.openById(SHEET_ID);
  const newSs = getMonthlySpreadsheet(year, month);

  const names = [
    `${year}年${month}月總表`,
    `${year}年${month}月瑜伽`,
    `${year}年${month}月羽球`,
    `${year}年${month}月網球`,
    `${year}年${month}月${REMAIN_SHEET_SUFFIX}`,
  ];

  const migrated = [];
  names.forEach(name => {
    const oldSheet = oldSs.getSheetByName(name);
    if (!oldSheet) return;
    if (newSs.getSheetByName(name)) return; // 新檔案已有同名分頁，跳過避免覆蓋

    const copied = oldSheet.copyTo(newSs);
    copied.setName(name);
    oldSs.deleteSheet(oldSheet);
    migrated.push(name);
  });

  removeDefaultSheet(newSs);
  Logger.log("搬遷完成：" + migrated.join("、"));
  return migrated;
}

function migrateOct2026() {
  migrateMonthToOwnFile(2026, 10);
}

function res(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function setCorsHeaders(output) {
  return output;
}
