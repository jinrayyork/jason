# Jason 工作追蹤

影音製作 KPI 線上同步儀表板。資料存在 Google 試算表，透過 Apps Script 網頁顯示。

- `index.html`：GitHub Pages 入口頁（主畫面圖示、使用說明），按鈕導向 Apps Script 網址
- `apps-script/Code.gs`、`apps-script/Index.html`：貼進試算表的 Apps Script 專案

## 更新儀表板程式
1. 修改 `apps-script/` 裡的檔案，貼回 Apps Script 編輯器並儲存
2. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）

## 更換儀表板網址
修改 `index.html` 最上方的 `TARGET`。

## 權限
- 管理者密碼、攝影師密碼存在 Apps Script「專案設定 → 指令碼屬性」（`ADMIN_PIN`、`EDITOR_PIN`），不在這個 repository 裡
- 試算表本身不要開編輯權限給攝影師
