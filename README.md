
# 🪑 Google Sheets 1:1 Interactive Seating Chart System
> 一套基於 Google Apps Script (GAS) 開發的 1:1 自動化動態座位表系統。專為學校班級管理設計，支援下拉選單智能檢索、即時數據同步與純學號顯示。

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?logo=googlesheets&logoColor=white)

---

## 📖 項目簡介 (Overview)

本項目專為簡化班主任與科任老師排座位而設計。透過 Google Apps Script 驅動，結合動態下拉選單（Data Validation）與 `VLOOKUP` 自動檢索公式，讓老師在選擇學生時能同時看到「學號與姓名」，選擇後格仔自動轉換為「純學號」，完美保持 1:1 教室佈局的整潔。

---

## ✨ 核心特色 (Features)

- 🏫 **1:1 教室排版還原**：忠實還原雙人桌連坐、走廊隔間、講台 (Teacher's Desk) 與置物柜 (Locker) 佈局。
- 🔍 **智能選單 + 自動簡化號碼**：選單顯示 `班號 - 中文名 英文名`（例如：`27 - 黃子謙 WONG TSZ HIM HUGO`），點選後自動觸發 `onEdit` 轉換為純學號 `27`。
- 🔄 **一鍵同步名單 (`Refresh`)**：老師在名單頁修改或新增學生資料後，撳一個掣即可自動重構下拉選單與關聯公式。
- ⚡ **自動化輔助功能**：提供「一鍵順序填入 1-28 號」與「一鍵清空座位」等快捷操作。
- 🎨 **純動態公式綁定**：座位表底層經由 `VLOOKUP` 與 `Student_List` 即時連動，改名無需重繪表格。

---

## 🛠️ 安裝與部署 (Installation)

只需 3 個步驟即可完成安裝：

1. **新建 Google Sheet**
   - 開啟或建立一張全新的 Google Sheets 試算表。

2. **開啟 Apps Script 編輯器**
   - 點擊頂部選單 **「擴充功能」 (Extensions)** ➔ **「Apps Script」**。
   - 將本專案的 `Code.gs` 程式碼完整複製並覆蓋至編輯器內。
   - 點擊上方 **儲存按鈕 💾**。

3. **初始化座位表**
   - 重新整理 Google Sheet 網頁，頂部選單會出現 **`座位表系統 🪑`**。
   - 點擊 **`座位表系統 🪑` ➔ `1. 重新建構 1:1 座位表`** 即可自動生成全套表格！

---

## 📂 系統架構 (System Architecture)

系統主要由兩個工作表 (Worksheets) 互動構成：


[ Google Sheets 試算表 ]
├── 📄 Student_List  (學生資料庫：儲存班別、學號、中英文姓名與選單顯示字串)
└── 📊 Seating_Chart (GUI 可視化座位表：連桌、走廊、講台與 Locker)

### 座位單元格結構 (Seat Matrix Architecture)

每張獨立桌子由三個垂直相連的單元格組成：


+-----------------------------------+  <- 頂層：Data Validation 下拉選單 (點選後存入純學號)
|               27                  |  <- 背景色 #f4f1ea，觸發 onEdit 事件
+-----------------------------------+  <- 中層：大字中文名
|             黃子謙                |  <- 公式：=IFERROR(VLOOKUP(頂層, Student_List!B:D, 3, FALSE), "")
+-----------------------------------+  <- 底層：英文全名
|      WONG TSZ HIM HUGO            |  <- 公式：=IFERROR(VLOOKUP(頂層, Student_List!B:C, 2, FALSE), "")
+-----------------------------------+

---

## 🎯 使用指南 (Usage)

### 1. 修改/新增學生名單
1. 切換至 `Student_List` 頁面，輸入同學的 Class, Class No, English Name 及 Chinese Name。
2. 點擊頂部選單 **`座位表系統 🪑` ➔ `2. 🔄 刷新名單與下拉選單`**（或點擊頁面上的同步按鈕）。

### 2. 排定座位
1. 切換至 `Seating_Chart` 頁面。
2. 點擊座位頂層格仔的下拉選單箭頭，選取對應學生。
3. 系統會自動將數值更換為純號碼，並顯示中英文名。

---

## 🤝 貢獻 (Contributing)

歡迎提交 Issue 或 Pull Request！如果你有任何優化建議或 bug 反饋，請隨時建立 Issue。

1. Fork 本專案
2. 建立你的 Feature 分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

---

## 📄 版權與許可證 (License & Copyright)

本專案採用 **MIT License** 授權。

```text
Copyright (c) 2026 Hugo Wong

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom it is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

Made with ❤️ by Hugo Wong

