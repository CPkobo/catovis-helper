# ディレクトリ情報をセットする
$workingdir = Get-Location
$macrodir = ".\myTest"

# ファイル名（定数）をセットする
$docmName = "$macrodir/CATOVIS_Helper.xlsm"
$files = Get-ChildItem $docmName

# 最新のファイルを取得する
# 最新のファイル（onGoing）がない場合、処理を終了する
$file = ""
if ($files.Length -lt 1) {
    Write-Host "${macrodir} にマクロ有効ファイルがありません"
    $file = $docmName
} else {
    $file = $files[0].FullName
}

Write-Host "最終ファイル： ${file} に処理を実行します。よろしいですか？"
$sw = Read-Host("[y/N]")
if ($sw -ne "y") {
    Write-Host "キャンセルされました"
    Read-host "Type any key:"
    Exit
}

# マクロをエクスポートしておく
Write-Host "マクロをエクスポートします"
cscript.exe .\vbac.wsf decombine /binary $macrodir /source ./code/

.\_commit-macro.ps1