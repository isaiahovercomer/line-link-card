# LINE網頁圖卡分享建立平台

把一個網址做成 LINE 圖卡發到群組：上方是網頁預覽圖，下方有「開啟網頁」「📲 加到手機主畫面」兩個按鈕，以及「轉傳給其他群組或朋友」連結。

- 貼上網址 → 自動帶入網頁的標題、說明、預覽圖（og:image），可再修改或上傳圖片
- 選圖卡色調 → 發送到 LINE 群組（用自己的 LINE 帳號，`shareTargetPicker`）
- 「我的圖卡」記錄發過的圖卡：發送次數、再發送、編輯、刪除；管理員看得到全部
- 「加到手機主畫面」：用手機瀏覽器開（LINE 裡面無法安裝），自己的網站直接帶 `?install=1`，其他網站先到本平台的安裝說明頁（`?add=網址`）

架構與「LINE簡易活動登記建立平台」相同：GitHub Pages 單頁 `index.html` ＋ LIFF ＋ Google Apps Script（`Code.gs`）＋ Google 試算表。

## 網址

| 用途 | 網址 |
|---|---|
| 平台 | https://liff.line.me/2012412693-IKxGZtWK |
| 預覽模式（假資料） | `index.html?demo&notour` |
| 加到主畫面說明頁 | `index.html?add=<網址>&t=<標題>&ic=<圖示>` |
| 轉傳頁 | `https://liff.line.me/<LIFF ID>?share=<圖卡代碼>` |
| 把平台裝到手機桌面 | `index.html?install=1` |

## 設定

1. 新的 Apps Script 專案貼上 `Code.gs`，執行 `setup`（會自動建立試算表「LINE網頁圖卡」），部署成網頁應用程式（執行身分：我；存取權：所有人）。
2. LINE Developers：Provider「CITC」的 LINE Login channel `2012412693` → LIFF → Add：Size Full、Endpoint 為 GitHub Pages 網址、Scopes `openid` `profile`、Share target picker ON。
3. `index.html` 最上方填 `LIFF_ID` 和 `GAS_URL`。
4. 管理員：把 LINE userId 填進「發布者」工作表。

## 自己的網站支援 `?install=1`

`OWN_HOSTS`（index.html）列出的網站，「加到手機主畫面」會直接開 `網址?install=1`，由網站自己跳出安裝提示（Android 安裝視窗、iPhone 圖解）。要加新網站，網站本身也要加上同樣的處理。
