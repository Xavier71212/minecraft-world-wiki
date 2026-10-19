# 更新網站：把本機「已發布」資料夾同步到網站儲存庫，並推送到 GitHub。
# GitHub 收到後會自動重建網站（約 1–2 分鐘）。
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Repo = Split-Path -Parent $PSScriptRoot
$ConfigPath = Join-Path $PSScriptRoot 'update-config.json'
$Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$WikiDir = [System.IO.Path]::GetFullPath((Join-Path $Repo $Config.wiki_folder))
$Published = Join-Path $WikiDir $Config.published_subfolder
$Content = Join-Path $Repo 'content'

function Say($msg, $color = 'Gray') { Write-Host $msg -ForegroundColor $color }

Say ''
Say '=== Minecraft World Wiki 網站更新 ===' 'Cyan'
Say "條目資料夾：$Published"

if (-not (Test-Path $Published)) {
    Say "找不到「已發布」資料夾：$Published" 'Red'
    Say '請確認 tools\update-config.json 裡的路徑。' 'Red'
    exit 1
}

# 找 git：先找系統安裝的，再找 GitHub Desktop 內附的
$Git = (Get-Command git -ErrorAction SilentlyContinue).Source
if (-not $Git) {
    $cand = Get-ChildItem "$env:LOCALAPPDATA\GitHubDesktop\app-*\resources\app\git\cmd\git.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1
    if ($cand) { $Git = $cand.FullName }
}
if (-not $Git) {
    Say '找不到 git。請先安裝並開啟一次 GitHub Desktop。' 'Red'
    exit 1
}
function G { & $Git -C $Repo @args }

# 1. 先取得網站上最新的版本（Claude 也可能更新過）
Say '取得最新版本……'
G -c core.quotepath=false pull --rebase --autostash --quiet 2>$null
if ($LASTEXITCODE -ne 0) { Say '無法從 GitHub 取得最新版本，繼續使用本機版本。' 'Yellow' }

# 2. 同步「已發布」→ content（只複製有變動的檔案，刪除已移除的檔案）
Say '同步條目……'
New-Item -ItemType Directory -Force -Path $Content | Out-Null
robocopy $Published $Content /MIR /XF '~$*' '.gitkeep' 'desktop.ini' 'Thumbs.db' /XD '_to_delete' /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { Say "同步失敗（robocopy 代碼 $LASTEXITCODE）" 'Red'; exit 1 }

# 3. 所有已寫好的條目名稱（未發布的條目在網站上顯示為紅字）
$names = Get-ChildItem $WikiDir -Recurse -Filter *.docx -File |
    Where-Object { $_.FullName -notmatch '\\_to_delete\\' -and $_.Name -notlike '~$*' } |
    ForEach-Object { $_.BaseName.Trim() } | Sort-Object -Unique
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines((Join-Path $Repo 'entries.txt'), [string[]]$names, $utf8)

# 4. 提交變更
G add -A content entries.txt
$changes = G -c core.quotepath=false status --porcelain -- content entries.txt
if (-not $changes) {
    Say ''
    Say '沒有新增或修改的條目，網站已是最新。' 'Green'
    exit 0
}
$list = $changes | Where-Object { $_ -match 'content/' } | ForEach-Object {
    $s = $_.Substring(0, 2).Trim(); $f = [System.IO.Path]::GetFileNameWithoutExtension($_.Substring(3).Trim('"'))
    switch -Regex ($s) { 'A|\?' { "新增 $f" } 'D' { "移除 $f" } default { "修改 $f" } }
}
Say ''
Say '這次的變更：' 'Cyan'
$list | ForEach-Object { Say "  $_" }

if (-not (G config user.name)) { G config user.name $Config.git_name }
if (-not (G config user.email)) { G config user.email $Config.git_email }
$msg = "更新條目 $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n`n" + ($list -join "`n")
G commit -q -m $msg
if ($LASTEXITCODE -ne 0) { Say '提交失敗。' 'Red'; exit 1 }

# 5. 推送到 GitHub
Say ''
Say '上傳到 GitHub……'
G push --quiet 2>$null
if ($LASTEXITCODE -ne 0) {
    Say '自動上傳沒有成功，已開啟 GitHub Desktop：請按右上角的「Push origin」完成上傳。' 'Yellow'
    $desk = "$env:LOCALAPPDATA\GitHubDesktop\GitHubDesktop.exe"
    if (Test-Path $desk) { Start-Process $desk }
    exit 0
}
Say ''
Say "完成！網站約 1–2 分鐘後更新：$($Config.site_url)" 'Green'
