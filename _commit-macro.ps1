$cm = Read-Host "変更をgitにcommitしますか？commitする場合はメッセージを入力してください"
if ($cm -eq "") {
} else {
    git add .
    git commit -m "macro: ${cm}"
}

Write-Host "すべての処理が終了しました"
Read-host "Type any key:"