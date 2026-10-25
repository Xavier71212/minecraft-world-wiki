# Minecraft World Wiki

以 Minecraft 世界為基礎、與現實世界平行的世界觀百科。

網站：https://xavier71212.github.io/minecraft-world-wiki/

## 運作方式

- `content/`：本機「Minecraft World Wiki\已發布」資料夾的鏡像。子資料夾就是網站上的分類。
- `content-en/`、`content-tora/`：本機「Minecraft World Wiki\English」「Minecraft World Wiki\托拉語」資料夾的鏡像，放英文版與托拉語版條目（與 `content/` 相同的子資料夾與檔名）。
- `entries.txt`：所有已寫好的條目名稱（含未發布），讓未發布條目在網站上顯示為紅字。
- `build.py`：把 Word 條目轉成網頁，輸出到 `_site/`。
- `config.json`：網站名稱、首頁特色條目、「你知道嗎？」等設定；`en`、`tora` 底下是英文版、托拉語版的對應設定與分類名稱。
- `static/`：樣式、搜尋程式、世界地圖；`static/fonts/` 是托拉語字母用的字型（DejaVu 的子集，授權見 LICENSE-DejaVu.txt）。
- `tools/docx_i18n.py`：Claude 翻譯條目用的工具（抽出 Word 條目文字、套回譯文並保留表格、圖片與連結）。
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

## 語言版本

- 網站有繁體中文（根目錄）、英文（`en/`）與托拉語 Ŧőřǎ（`tora/`）三種語言，每一頁右上角都有「語言」選單，可以切換到同一頁的另一種語言。
- 各語言的條目放在本機「Minecraft World Wiki」底下的「English」「托拉語」資料夾，子資料夾與檔名和「已發布」完全相同（檔名仍用中文條目名稱），譯名寫在文件的標題段落。
- 條目還沒有某種語言的版本時，語言選單中的該語言會顯示為灰色「尚無此語言版本」；譯文中連到尚未翻譯的條目時，會連到中文條目並標示「zh」。
- 譯文檔記錄了翻譯時中文原檔的版本；中文條目之後若有修改，該語言的頁面會提示「翻譯可能過時」，直到重新翻譯。
- 要再新增語言時：在 build.py 的 LANGS、LANG_ATTR、LANG_NAME、PREFIX、LANG_CONTENT、UI 加入該語言，在 config.json 加入對應設定，並在 tools/update-config.json 的 language_folders 加入本機資料夾。
