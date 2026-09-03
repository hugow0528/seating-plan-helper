/**
 * 1. 建立/更新 1:1 Google Sheets 座位表系統
 */
function setupSeatingSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // -------------------------------------------------------------
  // 1. 設定 Student_List
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
  listSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d9e1f2");
  
  // 建立並更新下拉選單數據範圍
  refreshStudentListAndDropdowns();

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
  
  var seatCols = [1, 2, 4, 5, 7, 8, 10];
  var defaultSeatNo = 1;
  
  for (var r = 0; r < 5; r++) {
    var startRow = 1 + (r * 4);
    
    for (var c = 0; c < seatCols.length; c++) {
      var col = seatCols[c];
      
      // Locker
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
      cellNo.setValue(defaultSeatNo <= 28 ? defaultSeatNo : "")
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
        
      // 黑框
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

  // 底部區域
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

  // 重新套用 Data Validation 下拉選單
  refreshStudentListAndDropdowns();
  onOpen();
}

/**
 * 2. 【核心】老師修改名單後，按下會自動更新選單及資料
 */
function refreshStudentListAndDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listSheet = ss.getSheetByName("Student_List");
  var chartSheet = ss.getSheetByName("Seating_Chart");
  
  if (!listSheet || !chartSheet) return;
  
  var lastRow = listSheet.getLastRow();
  if (lastRow < 2) lastRow = 2;
  
  // 1. 自動更新 Student_List 的 Column E 公式 (學號 - 中文名 英文名)
  var formulas = [];
  for (var i = 2; i <= lastRow; i++) {
    formulas.push([`=IF(B${i}="","", B${i} & " - " & D${i} & " " & C${i})`]);
  }
  listSheet.getRange(2, 5, lastRow - 1, 1).setFormulas(formulas);
  listSheet.autoResizeColumns(1, 5);
  
  // 2. 更新 Seating_Chart 上的選單 Data Validation 範圍
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(listSheet.getRange(`E2:E${lastRow}`), true)
    .setAllowInvalid(true)
    .build();
    
  var seatCols = [1, 2, 4, 5, 7, 8, 10];
  for (var r = 0; r < 5; r++) {
    var startRow = 1 + (r * 4);
    for (var c = 0; c < seatCols.length; c++) {
      var col = seatCols[c];
      if (col === 10 && r >= 2) continue; // 跳過 Locker
      
      chartSheet.getRange(startRow, col).setDataValidation(rule);
    }
  }
  
  SpreadsheetApp.getUi().alert("🔄 已成功同步名單！座位表選單與學生資料已全部更新。");
}

/**
 * 3. 選取完整選項後，自動截取並顯示純學號
 */
function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  
  if (sheet.getName() !== "Seating_Chart") return;
  
  var val = e.value;
  if (!val) return;
  
  if (typeof val === 'string' && val.indexOf(" - ") !== -1) {
    var parts = val.split(" - ");
    var num = parseInt(parts[0].trim(), 10);
    if (!isNaN(num)) {
      range.setValue(num);
    }
  }
}

/**
 * 4. 頂部選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('座位表系統 🪑')
    .addItem('1. 重新建構 1:1 座位表', 'setupSeatingSystem')
    .addItem('2. 🔄 刷新名單與下拉選單', 'refreshStudentListAndDropdowns')
    .addItem('3. 一鍵順序填入 1-28 號', 'autoFillSeats')
    .addItem('4. 清空所有座位', 'clearSeats')
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
