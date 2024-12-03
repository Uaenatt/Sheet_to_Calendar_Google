
# Google Calendar 自動化工具

此專案是一個 Google Apps Script 應用程式，用於自動化管理 Google Calendar 事件。

## 功能

- 自動將表單回應新增至 Google Calendar
- 支援借用器材和人力的不同事件類型
- 自動記錄和追蹤事件 ID
- 提供自訂選單進行操作

## 使用方式

1. 在 Google Sheets 中設置新試算表
2. 連結 Apps Script 專案
3. 設定您的 Calendar ID
4. 執行 `addCustomMenu` 函數來初始化

## 開發

使用 Google Apps Script 開發，主要使用：

- Google Calendar API
- Google Sheets API
- Apps Script UI Services

## 設定

初次使用時請執行 `firstUsed()` 函數並按照提示操作：

1. 設定您的 Calendar ID
2. 儲存程式碼
3. 執行 `addCustomMenu`

## 授權

此專案採用 MIT 授權條款。