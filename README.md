# Jason 工作追蹤

影音製作 KPI 線上同步儀表板。網頁放在 GitHub Pages，用公司 Google 帳號登入；資料存在 Google 試算表，由 Apps Script 驗證身分和讀寫。

- `index.html`：儀表板網頁（Google 登入、即時同步、加到主畫面）
- `apps-script/Code.gs`：貼進試算表 Apps Script 專案的後端

## 身分
在試算表「人員」分頁設定：

| 身分 | 可以做的事 |
|---|---|
| 管理者 | 全部：品牌配額、工時預算、明細、由配額重建明細 |
| 攝影師 | 新增、刪除、修改影片明細 |
| 檢視者 | 只能看 |

neverdie.com.tw、riseme.com.tw 的公司帳號即使不在名單上，登入後也是檢視者。改身分立即生效。

## 設定值
- `index.html` 最上方：`API_URL`（Apps Script 部署網址，結尾 `/exec`）、`CLIENT_ID`
- `apps-script/Code.gs` 最上方：`CLIENT_ID`、`ALLOWED_DOMAINS`、`APP_URL`

Google 登入用戶端 ID 與打卡系統共用（已授權來源 `https://jinrayyork.github.io`）。

## 更新後端
貼上新的 `Code.gs` 並儲存 → 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）。
