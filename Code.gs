/**
 * 1. 建立/更新 1:1 Google Sheets 座位表系統
 */
function setupSeatingSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // -------------------------------------------------------------
  // 1. 設定 Student_List (新增 Column E 作為有名單嘅選單)
  // -------------------------------------------------------------
  var listSheet = ss.getSheetByName("Student_List");
  if (!listSheet) {
    listSheet = ss.insertSheet("Student_List");
  } else {
    listSheet.clear();
  }
  
  var studentData = [
    ["Class", "Class No", "English Name", "Chinese Name", "Dropdown Display"],
    ["5B", 27, "WONG TSZ HIM HUGO", "黃子謙", ""]
  ];
  
  listSheet.getRange(1, 1, studentData.length, 5).setValues(studentData);
  
  // 設定 Column E 
  for (var i = 2; i <= studentData.length; i++) {
    listSheet.getRange(i, 5).setFormula(`=B${i} & " - " & D${i} & " " & C${i}`);
  }
  
  listSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d9e1f2");
  listSheet.autoResizeColumns(1, 5);

  // -------------------------------------------------------------
  // 2. 設定 Seating_Chart (GUI 座位表)
  // -------------------------------------------------------------
  var chartSheet = ss.getSheetByName("Seating_Chart");
  if (!chartSheet) {
    chartSheet = ss.insertSheet("Seating_Chart");
  } else {
    chartSheet.clear();
    chartSheet.clearFormats();
  }
  
  chartSheet.setHiddenGridlines(false);
  
  // validation 指向 Student_List!E2:E36（帶名單選項），並允許輸入轉換後的數字
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(listSheet.getRange("E2:E36"), true)
    .setAllowInvalid(true)
    .build();

  var seatCols = [1, 2, 4, 5, 7, 8, 10];
  var defaultSeatNo = 1;
  
  for (var r = 0; r < 5; r++) {
    var startRow = 1 + (r * 4);
    
    for (var c = 0; c < seatCols.length; c++) {
      var col = seatCols[c];
      
      // J 欄第 3 排開始係 Locker
      if (col === 10 && r >= 2) {
        var lockerRange = chartSheet.getRange(startRow, col, 3, 1);
        lockerRange.merge()
          .setValue("Locker")
          .setFontSize(13)
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setBackground("#e8e8e8")
          .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        continue;
      }
      
      // 桌子頂層：班號選單
      var cellNo = chartSheet.getRange(startRow, col);
      cellNo.setDataValidation(rule)
        .setValue(defaultSeatNo <= 28 ? defaultSeatNo : "")
        .setFontSize(11)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBackground("#f4f1ea");
        
      // 桌子中層：中文名
      var cellCname = chartSheet.getRange(startRow + 1, col);
      cellCname.setFormula(`=IFERROR(VLOOKUP(${cellNo.getA1Notation()}, Student_List!B:D, 3, FALSE), "")`)
        .setFontSize(14)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
        
      // 桌子底層：英文名
      var cellEname = chartSheet.getRange(startRow + 2, col);
      cellEname.setFormula(`=IFERROR(VLOOKUP(${cellNo.getA1Notation()}, Student_List!B:C, 2, FALSE), "")`)
        .setFontSize(9)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
        
      // 單張桌子黑框
      chartSheet.getRange(startRow, col, 3, 1)
        .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        
      defaultSeatNo++;
    }
  }

  // 行高設定
  chartSheet.setRowHeight(4, 12);
  chartSheet.setRowHeight(8, 12);
  chartSheet.setRowHeight(12, 12);
  chartSheet.setRowHeight(16, 12);
  chartSheet.setRowHeight(20, 16);

  // -------------------------------------------------------------
  // 3. 底部區域 (Monitor, Teacher's Desk, Form Teacher, Locker)
  // -------------------------------------------------------------
  var bRow = 21;
  chartSheet.getRange(bRow, 1).setValue("Monitor:").setFontWeight("bold").setFontSize(10);
  chartSheet.getRange(bRow + 1, 1).setValue("Monitress:").setFontWeight("bold").setFontSize(10);
  chartSheet.getRange(bRow, 2, 2, 1).setBorder(false, false, true, false, false, false);

  var tDesk = chartSheet.getRange(bRow, 4, 3, 2);
  tDesk.merge()
    .setValue("Teacher's Desk\n(5B)\nSep 2026")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#fafafa")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  chartSheet.getRange(bRow, 7).setValue("Form Teacher:").setFontWeight("bold").setFontSize(10);
  chartSheet.getRange(bRow, 8, 2, 1).setBorder(false, false, true, false, false, false);

  var bLocker = chartSheet.getRange(bRow, 10, 3, 1);
  bLocker.merge()
    .setValue("Locker")
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#e8e8e8")
    .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 欄寬設定
  var colWidths = { 1: 125, 2: 125, 3: 18, 4: 125, 5: 125, 6: 18, 7: 125, 8: 125, 9: 18, 10: 130 };
  for (var colIdx in colWidths) {
    chartSheet.setColumnWidth(parseInt(colIdx), colWidths[colIdx]);
  }

  onOpen();
  SpreadsheetApp.getUi().alert("✅ 座位表已升級！選單現已顯示「學號 + 姓名」，選擇後會自動顯示為純學號！");
}

/**
 * 2. 自動監聽事件 (選取後自動切換為純學號)
 */
function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== "Seating_Chart") return;
  
  var val = e.value;
  if (!val) return;
  
  // 只要字串包含 " - "（如 "27 - 黃子謙 WONG TSZ HIM HUGO"），就自動切換回前面的數字 "27"
  if (typeof val === 'string' && val.indexOf(" - ") !== -1) {
    var parts = val.split(" - ");
    var num = parseInt(parts[0].trim(), 10);
    if (!isNaN(num)) {
      range.setValue(num);
    }
  }
}

/**
 * 頂部工具列選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('座位表系統 🪑')
    .addItem('1. 重新建構 1:1 座位表 (含名字下拉選單)', 'setupSeatingSystem')
    .addItem('2. 一鍵順序填入 1-28 號', 'autoFillSeats')
    .addItem('3. 清空所有座位', 'clearSeats')
    .addToUi();
}

function autoFillSeats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheet = ss.getSheetByName("Seating_Chart");
  if (!chartSheet) return;
  
  var seatCols = [1, 2, 4, 5, 7, 8, 10];
  var seatNo = 1;
  
  for (var r = 0; r < 5; r++) {
    var startRow = 1 + (r * 4);
    for (var c = 0; c < seatCols.length; c++) {
      var col = seatCols[c];
      if (col === 10 && r >= 2) continue;
      
      chartSheet.getRange(startRow, col).setValue(seatNo);
      seatNo++;
      if (seatNo > 28) break;
    }
    if (seatNo > 28) break;
  }
}

function clearSeats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheet = ss.getSheetByName("Seating_Chart");
  if (!chartSheet) return;
  
  var seatCols = [1, 2, 4, 5, 7, 8, 10];
  for (var r = 0; r < 5; r++) {
    var startRow = 1 + (r * 4);
    for (var c = 0; c < seatCols.length; c++) {
      var col = seatCols[c];
      if (col === 10 && r >= 2) continue;
      chartSheet.getRange(startRow, col).clearContent();
    }
  }
}
