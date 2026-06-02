# 台灣雅芳每日即時戰報網站

這個資料夾是一個可放到固定網址的靜態網站。員工每天只需要打開同一個網址，看到的數字由 `war_data.js` 更新。

## 檔案用途

- `index.html`: 戰報網頁入口。放到網站後，員工開網址就會看到這頁。
- `war_data.js`: 每天更新的即時資料檔。
- `update-war-data.ps1`: 從最新 Excel 產生 `war_data.js` 的一鍵更新腳本。

## 每日更新流程

1. 將當天最新的 `real time .xlsx` 放進這個資料夾，或放在桌面。
2. 右鍵執行 `update-war-data.ps1`。
3. 確認網頁打開後數字正確。
4. 將更新後的 `war_data.js` 上傳或同步到固定網址的網站空間。

如果網站空間不是自動同步，也可以連同 `index.html` 一起上傳；但通常每天只需要更新 `war_data.js`。

## 固定網址建議

推薦用靜態網站服務，例如：

- Cloudflare Pages
- GitHub Pages
- 公司內部網站或 NAS
- SharePoint/OneDrive，如果公司環境允許直接載入 HTML 與 JS

最理想的員工體驗是固定網址，例如：

`https://你的網址/`

每天只更新 `war_data.js`，網址不變。

## 注意事項

- `index.html` 需要和 `war_data.js` 放在同一層資料夾。
- 更新腳本需要 Windows 電腦上可使用 Excel。
- 如果員工看不到最新數字，請先重新整理瀏覽器；必要時用 Ctrl + F5 強制刷新。
