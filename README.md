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

## 開發步驟

### 環境設定
1. 安裝 clasp CLI 工具
```bash
npm install @google/clasp -g
```

2. 確認 clasp 版本
```bash
clasp -v
```

3. 登入 Google 帳戶
```bash
clasp login
```

4. 複製 host URL 並在新終端機執行
```bash
curl <host-URL>
```

### 開發流程

#### 測試環境

1. 推送程式碼到測試環境
```bash
clasp push
```

#### 部署到線上環境
1. 同步測試環境程式碼到線上環境
```bash
rsync -av --delete /workspaces/Sheet_to_Calendar_Google/test/src/ /workspaces/Sheet_to_Calendar_Google/online/src/
```

2. 推送程式碼到線上環境
```bash
clasp push
```

### 注意事項
- 確保在推送前已完成所有測試
- 保持測試環境和線上環境的 .clasp.json 設定正確
- 每次部署前請確認程式碼變更的影響範圍