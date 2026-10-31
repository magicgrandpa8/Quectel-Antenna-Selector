# ============================================================================
#  客戶分享資料夾：為「每一份」datasheet 建立「任何人可檢視」連結，並以未登入身分逐一驗證
#  1) 以你的 Microsoft 365 帳號登入 (Microsoft Graph)
#  2) 列出客戶分享資料夾中的所有 PDF
#  3) 每份 PDF 建立 type=view / scope=anonymous 連結 (已存在則沿用同一條)
#  4) 不帶任何登入資訊 (全新工作階段) 開啟每條連結，確認不會被導向登入頁
#  5) 輸出 datasheet_links_all.csv 到桌面「Datasheet_links_for_Claude」資料夾，上傳給 Claude
# ============================================================================
$ErrorActionPreference = 'Stop'
$Folder  = 'Customer/Antenna FAE/0.For Customer/Quectel_Antenna_Selector/V3.7_20260916_Rev26.Q3'
$OutDir  = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Datasheet_links_for_Claude'
$OutCsv  = Join-Path $OutDir 'datasheet_links_all.csv'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

function Encode-Path([string]$p) { ($p -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/' }

# 未登入驗證：全新工作階段 (沒有 cookie、沒有帳號) 開啟連結；最後停在登入頁 = 不是公開連結
function Test-Anonymous([string]$url) {
    try {
        $s = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $r = Invoke-WebRequest -Uri $url -WebSession $s -UseBasicParsing -MaximumRedirection 10 -TimeoutSec 60
        $final = if ($r.BaseResponse.ResponseUri) { $r.BaseResponse.ResponseUri.AbsoluteUri } else { $r.BaseResponse.RequestMessage.RequestUri.AbsoluteUri }
        if ($final -match 'login\.microsoftonline\.com|_layouts/15/Authenticate|AccessDenied|guestaccess\.aspx\?.*error') { return @('NG', 'redirected to sign-in') }
        if ($r.StatusCode -eq 200) { return @('OK', '') }
        return @('NG', "HTTP $($r.StatusCode)")
    } catch { return @('NG', $_.Exception.Message) }
}

Write-Host ''
Write-Host '=== 建立並驗證「任何人可檢視」的 datasheet 連結 ===' -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host '第一次執行：安裝 Microsoft.Graph.Authentication 模組 (約 1 分鐘)…'
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication
Write-Host '請在跳出的視窗中，用公司帳號登入並同意權限…'
Connect-MgGraph -Scopes 'Files.ReadWrite' -NoWelcome

# 列出資料夾中所有 PDF (分頁)
$files = @()
$uri = "https://graph.microsoft.com/v1.0/me/drive/root:/$(Encode-Path $Folder):/children?`$top=200&`$select=id,name,file"
while ($uri) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $uri
    $files += $page.value | Where-Object { $_.file -and $_.name -match '\.pdf$' }
    $uri = $page.'@odata.nextLink'
}
Write-Host "資料夾中共有 $($files.Count) 份 PDF"

$rows = @(); $n = 0; $ok = 0
foreach ($f in ($files | Sort-Object name)) {
    $n++
    $url = ''; $status = 'NG'; $note = ''
    try {
        $body = @{ type = 'view'; scope = 'anonymous' } | ConvertTo-Json
        $r = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/drive/items/$($f.id)/createLink" -Body $body -ContentType 'application/json'
        $url = $r.link.webUrl
        $chk = Test-Anonymous $url
        $status = $chk[0]; $note = $chk[1]
    } catch { $note = "建立連結失敗：$($_.Exception.Message)" }
    if ($status -eq 'OK') { $ok++ }
    $rows += [pscustomobject]@{ Datasheet_File = $f.name; OneDrive_URL = $url; Anonymous_Check = $status; Note = $note }
    Write-Progress -Activity '建立並驗證連結' -Status "$n / $($files.Count)  $($f.name)" -PercentComplete ($n * 100 / [math]::Max(1, $files.Count))
    if ($status -ne 'OK') { Write-Host "  [NG] $($f.name)：$note" -ForegroundColor Red }
}
Disconnect-MgGraph | Out-Null
$rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host "完成：$($files.Count) 份 PDF，未登入可開啟 $ok 份，失敗 $($files.Count - $ok) 份" -ForegroundColor Cyan
if ($ok -lt $files.Count) { Write-Host '失敗的檔案已在上方列出 (CSV 的 Anonymous_Check = NG)。' -ForegroundColor Yellow }
Write-Host "結果：$OutCsv"
Write-Host '請把這個 CSV 上傳給 Claude。'
Start-Process explorer.exe $OutDir
Read-Host '按 Enter 結束'
