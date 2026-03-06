# 記帳與財務預測

這是一個純前端的記帳與財務預測工具，可直接部署到 GitHub Pages。網站資料只保存在使用者各自瀏覽器的 `localStorage` 中，不會上傳到伺服器。

## 功能

- 設定目前存款與預測月數
- 管理每月或每年固定月份的定期收入與支出
- 管理分期付款
- 匯出與匯入正式 JSON 備份
- 顯示未來月度預測表格與餘額趨勢圖

## 資料行為

- 每位使用者的資料只存在自己使用該網站的瀏覽器
- 不同使用者彼此不共享資料
- 清除瀏覽器資料後，本機內容會消失
- 若要跨裝置搬移資料，請使用 JSON 備份匯出與匯入

## GitHub Pages 部署

1. 建立新的 GitHub repository。
2. 將本專案內容推到該 repository 的 `master` branch root。
3. 在 repository 內建立 `.github/workflows/pages.yml`，使用 GitHub Actions 部署 GitHub Pages。
4. 到 GitHub repository 的 `Settings > Pages`。
5. 在 `Build and deployment` 的 `Source` 選擇 `GitHub Actions`。
6. push 到 `master` 後等待 Actions 執行與 Pages 發佈完成。

部署完成後，網址格式通常為：

`https://<username>.github.io/<repo-name>/`

## 注意事項

- 這是公開網址，不是受控存取。知道連結的人都能開啟網站。
- 專案已加入 `noindex` 與 `robots.txt` 以降低被搜尋引擎收錄的機率，但這不是安全保護。
- 圖表依賴 CDN 載入 Chart.js，若網路阻擋 CDN，圖表可能無法顯示。
