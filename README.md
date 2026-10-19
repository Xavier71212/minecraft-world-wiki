# Minecraft World Wiki

以 Minecraft 世界為基礎、與現實世界平行的世界觀百科。

網站：https://xavier71212.github.io/minecraft-world-wiki/

## 運作方式

- `content/`：本機「Minecraft World Wiki\已發布」資料夾的鏡像。子資料夾就是網站上的分類。
- `entries.txt`：所有已寫好的條目名稱（含未發布），讓未發布條目在網站上顯示為紅字。
- `build.py`：把 Word 條目轉成網頁，輸出到 `_site/`。
- `config.json`：網站名稱、首頁特色條目、「你知道嗎？」等設定。
- `static/`：樣式、搜尋程式、世界地圖。
- `.github/workflows/deploy.yml`：每次推送後自動建站並發布到 GitHub Pages。

## 更新網站

**方法一：雙擊「更新網站.bat」**
同步「已發布」資料夾、提交並推送；GitHub 約 1–2 分鐘後完成重建。需要先安裝並登入 GitHub Desktop。

**方法二：在 Claude 的「虛擬世界觀」專案中說「更新網站」**
Claude 會比對本機資料夾與這個儲存庫，只上傳新增、修改或刪除的條目。

## 條目規則

- 一個條目一個 Word 檔，檔名就是條目名稱。
- 檔案開頭的表格會成為右側資訊框；表格中的圖片（旗幟、標誌）依原檔方向顯示，不旋轉。
- 條目中指向其他條目 Word 檔的超連結，會依目標狀態自動顯示：已發布為藍色連結，已寫好但未發布為紅字，其他為一般文字。
