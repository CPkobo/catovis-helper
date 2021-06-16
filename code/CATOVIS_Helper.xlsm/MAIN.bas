Attribute VB_Name = "MAIN"
Option Explicit


Dim ws As Worksheet
Dim ready As Boolean
Dim makebu As Boolean
Dim h As CHelper
Dim req As Cat_Request

'フォームを呼び出した際に、ワークシートにフォーカスが残るようにするためのライブラリ
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

'CHelperクラスがなければ初期化する
Public Sub standby()

    If ready = False Then

        Set h = New CHelper
        Set req = New Cat_Request
        ready = True
        
    End If

End Sub

'バックアップの設定を変更する
'＃TODO アドインタブからも変更できるようにする
Public Sub set_backup()

    If ThisWorkbook.Sheets("PREFERENCE").Range("E6").Value = 1 Then
        makebu = True
    
    Else
        makebu = False
        
    End If

End Sub

'マクロのショートカットを設定する
Public Sub apply_shortcut()

    Set ws = ThisWorkbook.Sheets("PREFERENCE")

    Application.MacroOptions macro:="lint_merge", ShortcutKey:=ws.Range("C2").text
    Application.MacroOptions macro:="lint_split", ShortcutKey:=ws.Range("C3").text
    Application.MacroOptions macro:="lint_insert", ShortcutKey:=ws.Range("C4").text
    Application.MacroOptions macro:="lint_adjust", ShortcutKey:=ws.Range("C5").text

End Sub

'ヘルパーフォームを呼び出す
Public Sub call_helper()
    
    If CatovisH.inculedFiles.ListCount = 0 Then
        Call set_files_inForm
    End If

    CatovisH.Show

End Sub

'ヘルパーフォームのファイルリストを更新
Private Sub set_files_inForm()

    Dim list As Variant
    Dim i As Integer

    'ヘルパークラスから格納しているファイルリストを取得
    list = h.files_list
    
    For i = 0 To UBound(list)
    
        CatovisH.inculedFiles.AddItem (list(i))
        CatovisH.inculedFiles2.AddItem (list(i))
    
    Next i

End Sub

'ALIGNシートの一行目にヘッダーを書き込む
Public Sub write_header()

    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    ws.Range("A1").Value = "原文"
    ws.Range("B1").Value = "訳文"
    ws.Range("C1").Value = "同じ"
    ws.Range("D1").Value = "重複"

End Sub

'CATOVISのアラインファイルを読み込む
'フォームのSTARTボタンか、アドインタブの「開く」から実行することを想定
Public Sub open_align_file()

    Call standby

    Dim strFiles As String, files As String
    Dim i As Integer

    'ファイルの読み込みダイアログ　ここから
    With Application.FileDialog(msoFileDialogFilePicker)
        'ファイルの複数選択を可能にする
        .AllowMultiSelect = False
        'ファイルフィルタのクリア
        .Filters.Clear
        'ファイルフィルタの追加
        .Filters.Add "その他", "*.tsv"
        '初期表示フォルダの設定
        .InitialFileName = ThisWorkbook.Path

        If .Show = -1 Then  'ファイルダイアログ表示
            ' [ OK ] ボタンが押された場合
            For i = 1 To .SelectedItems.Count
                strFiles = .SelectedItems(i)
            Next i

        Else
            ' [ キャンセル ] ボタンが押された場合
            MsgBox "ファイル選択がキャンセルされました。", vbExclamation
        End If
        
    End With
    
    'ファイルの読み込みダイアログ　ここまで

    '読み込んだ文字列を展開する
    If strFiles <> "" Then

        Application.ScreenUpdating = False

        '既存のデータを消して、書式を文字列にする
        Call start_preparation
        
        '実際のtsv→Excel変換処理
        Call readin_TSV(strFiles)
    
        '読み込んだデータを人の目に見やすくする
        Call decorate_friendly
        
        'シートを複製しておく
        Call backup_align_sheet
        
        'フォームのファイルリストを更新する
        Call set_files_inForm
        
        Application.ScreenUpdating = True
        
        MsgBox ("読み込みが完了しました")
        
    End If
    
    'フォーカスをワークシートに戻す
    SetFocus Application.hwnd

End Sub

'ファイルの初期設定をする
'具体的には以下の操作
'ファイルリストの削除、データの削除、ヘッダーの追記、フィルターのセット、書式のセット
'新規ファイルの読み込み時に自動で実行される
'アップロード前に手動で実行することも可
Public Sub start_preparation()

    Call standby

    CatovisH.inculedFiles.Clear
    CatovisH.inculedFiles2.Clear

    With ThisWorkbook.Sheets("ALIGN")
        .Select
        .Range(Cells(2, 1), Cells(h.end_row, 1)).EntireRow.Delete
        
        With .Range("A:E")
            .Interior.pattern = xlNone
            .NumberFormatLocal = "@"
        End With
        
        'ヘッダーを書き加える
        Call write_header
        
        If .AutoFilterMode = True Then
            .Range("A:D").AutoFilter
            .Range("A:D").AutoFilter
                
        Else
            .Range("A:D").AutoFilter
            
        End If
        
        .Range("A2").Select
        
    End With
    
    With ThisWorkbook.Sheets("STATUS")
    
        .Range("B:B").ClearContents
    
    End With

End Sub

'指定されたTSVファイルから文字列を抽出してセルに代入する
Private Sub readin_TSV(ByVal target As String)
    Dim buf As String
    Dim lines() As String, lineBuf() As String
    Dim line As Variant, data As Variant
    Dim i As Long
    
    Dim innerFiles As String
    
    ReDim rngdata(99999, 1)
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile target
        buf = .ReadText
        .Close
    
    End With
    
    lines = Split(buf, vbLf)
    
    i = 1
    rngdata(0, 0) = "原文"
    rngdata(0, 1) = "訳文"
    
    For Each line In lines
        If line <> "" Then
            lineBuf = Split(line, vbTab)
            If UBound(lineBuf) = 1 Then
                rngdata(i, 0) = lineBuf(0)
                rngdata(i, 1) = lineBuf(1)
            ElseIf UBound(lineBuf) = 0 Then
                If lineBuf(0) <> "" Then
                    rngdata(i, 0) = lineBuf(0)
                End If
            End If
            
            If h.fMark.test(rngdata(i, 0)) Then
                '読み込んだ行がファイル名の場合、ファイルリストとして別保管する
                innerFiles = innerFiles + vbCrLf + Replace(rngdata(i, 0), "_@@_ ", "")
            
            End If
        
            i = i + 1
        
        End If
    
    Next line
    
    '改行記号が先頭に入ってしまうので削除する
    innerFiles = Replace(innerFiles, vbCrLf, "", 1, 1)
    
    'ヘルパークラスのプロパティをセットする
    Call h.set_status(target, innerFiles, i)
    'ステータスのシートを更新
    Call h.write_status
    'アドインタブのファイルリストを更新
    'Call set_files_inBar(innerFiles)
    
    ws.Range("A1").Resize(i, 2) = rngdata
    
End Sub

'区切り文字に色をつけていく
Private Sub decorate_friendly()

    Dim a As String
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    ws.Activate
 
 
    ' ファイル名ともう一段下位レベルの区切り記号に色を付ける
    'Word
    If WorksheetFunction.CountIf(ws.Range("A:A"), WORD_PARA_MARK) > 0 Then
        Call filter_and_color(WORD_FILE_MARK, 37)
        Call filter_and_color(WORD_PARA_MARK, 39)
        
        If WorksheetFunction.CountIf(ws.Range("A:A"), WORD_TBOX_MARK) > 0 Then
            Call filter_and_color(WORD_TBOX_MARK, 39)
        
        End If
        
        If WorksheetFunction.CountIf(ws.Range("A:A"), WORD_TABLE_MARK) > 0 Then
            Call filter_and_color(WORD_TABLE_MARK, 39)
        
        End If
        
    End If
    
    'Excel
    If WorksheetFunction.CountIf(ws.Range("A:A"), "_@λ_ SHEET1 _λ@_") > 0 Then
        Call filter_and_color(EXCEL_FILE_MARK, 50)
        Call filter_and_color(EXCEL_SHEET_MARK, 43)
    End If
    
    'PPT
    If WorksheetFunction.CountIf(ws.Range("A:A"), "_@λ_ SLIDE1 _λ@_") > 0 Then
        Call filter_and_color(PPT_FILE_MARK, 46)
        Call filter_and_color(PPT_SLIDE_MARK, 45)
    End If
    
    
    ' EOF
    Call filter_and_color(FILE_END_MARK, 48)
    
    '最後にフィルタ解除
    ws.ShowAllData
    ws.Range("A2").Select

End Sub

'フィルターの条件と色を受け取って反映するための内部プロシージャ
Private Sub filter_and_color(ByVal cr As String, ByVal cx As Long)

    Columns("A:A").Select
    Selection.AutoFilter Field:=1, _
        Criteria1:="=" & cr
    
    ws.Range(Cells(2, 1), Cells(h.end_row, 2)).Select
        
    Selection.Interior.ColorIndex = cx

End Sub

'ファイル名を受け取ってフォーカスを移動するサブプロシージャ
'検索機能を利用している
Public Sub move_to_file(ByVal name As String)

    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    Dim srcCol As Range
    Dim gotRow As Integer
    Dim nameWithMark As String
    
    Application.ScreenUpdating = False
    
    nameWithMark = "_@@_ " + name

    Set srcCol = ws.Columns("A:A")
    
    '念のため「セルの完全一致」を使用する
    If srcCol.Find(what:=nameWithMark, LookAt:=xlWhole) Is Nothing Then
        Exit Sub
    
    Else
        gotRow = srcCol.Find(what:=nameWithMark, LookAt:=xlWhole).Row
    
    End If
    
    'セルの完全一致を解除しておく
    srcCol.Find(what:="", LookAt:=xlPart).Activate
    
    Application.ScreenUpdating = True
    
    ws.Cells(gotRow, 1).Select

End Sub

'セルの結合処理
Public Sub lint_merge()
Attribute lint_merge.VB_ProcData.VB_Invoke_Func = "M\n14"

    Call standby

    Dim strInRange As String, strMerged As String
    Dim colsNum As Integer, rowsNum As Integer, i As Integer, j As Integer
    Dim crtRng As Range
    
    Set crtRng = Selection
    
    If crtRng.Rows.Count = 1 Then
        MsgBox ("2行以上選択してください")
        Exit Sub
    End If
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    colsNum = crtRng.Columns.Count
    rowsNum = crtRng.Rows.Count
    
    For i = 1 To colsNum
        strMerged = ""
        
        For j = 1 To rowsNum
            
            strInRange = crtRng(j, i).text
            
            '文字列型でなければ文字列型に変換して追加代入
            If VarType(strInRange) <> 8 Then
                strMerged = strMerged + Str(strInRange)
            
            Else
                '区切り文字は結合されないようにする
                If h.fMark.test(strInRange) = False And h.sMark.test(strInRange) = False Then
                    If strInRange <> FILE_END_MARK Then
                        strMerged = strMerged + strInRange
                    
                    Else
                        Exit For
                        
                    End If
                
                Else
                    Exit For
                
                End If
        
            End If
            
        Next j
        
        crtRng(1, i).Select
        
        '結合処理
        Selection(1, 1).Value = strMerged
        'Selection.Resize(rowsNum - 1, 1).Select
        If j > 2 Then
            Selection.Resize(j - 2, 1).Select
            Selection.Offset(1, 0).Select
            Selection.Delete Shift:=xlShiftUp
        End If
        
    Next i

End Sub

'選択範囲内のテキストをデリミタに基づいて分割
'分割後のセグメント数に基づいて複数のセルに代入
'#TODO デリミタの複数対応
Public Sub lint_split()
Attribute lint_split.VB_ProcData.VB_Invoke_Func = "S\n14"

    Call standby

    Dim allText As String, singleText As String, delimiter As String
    Dim sentences() As String
    Dim sentenceNum As Long
    
    Dim isEndWithDelim As Boolean
    
    Dim crtRng As Range
    Dim singleRng As Range
    Dim rngNum As Long
    Dim i As Integer, j As Integer
    
    '分割禁止情報のチェック
    Set crtRng = Selection
    
    '分割は1列のみに対して有効
    If crtRng.Columns.Count > 1 Then
        MsgBox "列は1列だけを選択してください"
        End
        
    End If
    
    '選択範囲の行数を数える
    rngNum = crtRng.Count
    
    '分割は20行まで
    If rngNum > 1000 Then
        MsgBox "行の選択は20行までにしてください"
        End
        
    End If
    
    'デリミタを設定する
    If crtRng.Column = 1 Then
        delimiter = ThisWorkbook.Sheets("PREFERENCE").Range("E2").text
    
    ElseIf crtRng.Column = 2 Then
        delimiter = ThisWorkbook.Sheets("PREFERENCE").Range("E3").text
        
    Else
        MsgBox "分割機能は原文列 または 訳文列でのみ有効です"
        End
    
    End If
    
    '禁止設定の確認ここまで
    
    'バックアップを行う
    If makebu Then
        Call backup_align_sheet
    End If
    
    '一つの変数に一旦すべての文字列を格納する
    allText = ""
    
    For j = 1 To rngNum
        
        singleText = crtRng(j).text
        
        ' 区切り記号に当たったら格納を中断する
        If h.sMark.test(singleText) Then
            Exit For
            
        ElseIf h.fMark.test(singleText) Then
            Exit For
            
        ElseIf singleText = FILE_END_MARK Then
            Exit For
        
        End If
        
        allText = allText + singleText
    
    Next j
    
    j = j - 1
    
    '最後がデリミタかどうかの判定をしておく
    isEndWithDelim = Right(allText, 1) = delimiter
    
    'デリミタで分割する
    sentences = Split(allText, delimiter)
    sentenceNum = UBound(sentences)
    
    '分割の必要がなかった場合は処理を終了
    If sentenceNum = 0 Then
        Exit Sub
        
    End If
    
    'Splitで消えたデリミタを補う
    For i = 0 To sentenceNum - 1
        
        sentences(i) = sentences(i) + delimiter
        
    Next i
    
    '最後のセグメントがデリミタで終わっていた場合、空の文字列が入っているので
    'sentenceNumをデクリメント
    
    If isEndWithDelim Then
        sentenceNum = sentenceNum - 1
        
    End If
    
    
    '分割後のセグメント数が選択範囲より少ない場合、
    '分割結果をセルに記入したのち、不要な文を削除する
    'If rngNum >= sentenceNum + 1 Then
    If j >= sentenceNum + 1 Then
        For i = 0 To sentenceNum
            Selection(i + 1).Value = sentences(i)
        Next i
        
        '最後のセグメント以降は内容を削除
        'For i = sentenceNum + 1 To rngNum - 1
        For i = sentenceNum + 1 To j - 1
            Selection(i + 1).Clear
        Next i
        
    '分割後のセグメント数が選択範囲より多い場合、
    '選択範囲以降はセルを挿入しながら書き込む
    Else
        'For i = 1 To rngNum
        For i = 1 To j
            Selection(i).Value = sentences(i - 1)
        Next i
        
        'Selection(rngNum).Select
        Selection(j).Select
        
        'For i = rngNum To sentenceNum
        For i = j To sentenceNum
            Selection.Offset(1, 0).Select
            Selection.Insert (xlShiftDown)
            Selection.Value = sentences(i)
        Next i
    End If
    
    'Call adjust_to_next_mark(crtRng.Row)
    'Call blank_row_delete_to_next(crtRng.Row)

End Sub

' クリップボードの中身を押し出して挿入する
Public Sub lint_insert()
Attribute lint_insert.VB_ProcData.VB_Invoke_Func = "I\n14"
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    Selection.Insert Shift:=xlDown
    ActiveSheet.Paste

End Sub

' 区切り記号に基づく調整
'調整は原文に訳文を合わせる形で行う
Public Sub lint_adjust()
Attribute lint_adjust.VB_ProcData.VB_Invoke_Func = "A\n14"

    Call standby

    Dim i As Integer, j As Integer, k As Integer
    Dim srcSectDist As Integer, tgtSectDist As Integer
    Dim crtRow As Integer, section_row As Integer, start_row As Integer
    Dim crtSrc As String, crtTgt As String
    Dim sIsFile As Boolean, sIsSection As Boolean, sIsEOF As Boolean
    Dim tIsFile As Boolean, tIsSection As Boolean, tIsEOF As Boolean
    Dim shouldAdjust As Boolean, toAdjust As Integer, toAdujstStr As String
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    ' 現在の行数を取得
    crtRow = ActiveCell.Row
    
    crtSrc = ws.Cells(crtRow, 1).text
    crtTgt = ws.Cells(crtRow, 2).text
        
    If makebu Then
        Call backup_align_sheet
    End If
    
    section_row = adjust_to_previous_mark(crtRow)
    start_row = section_row + 1
    
    Do While start_row - crtRow < h.modify_limit
    
        start_row = adjust_to_next_mark(start_row) + 1
        
    Loop
    
    Call blank_row_delete(section_row)
    
    ws.Cells(crtRow, 1).Select

End Sub

' 直近の上側の区切りが一致しているかのテスト
' 原文側、訳文側それぞれセルの距離を測り、差異のある分だけ調整する
'調整が終わったら調整後の行数を返す
Private Function adjust_to_previous_mark(ByVal crtRow As Integer) As Integer
    
    Dim i As Integer
    Dim crtSrc As String, crtTgt As String
    Dim srcFoundSect As Boolean, tgtFoundSect As Boolean
    Dim srcSectDist As Integer, tgtSectDist As Integer
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    For i = 1 To crtRow - 1
        
        If srcFoundSect = False Then
            crtSrc = ws.Cells(crtRow - i, 1).text
            
            If h.sMark.test(crtSrc) Then
                srcSectDist = i
                srcFoundSect = True
                
                If tgtFoundSect Then
                    Exit For
                    
                End If
            End If
        End If
        
        If tgtFoundSect = False Then
            crtTgt = ws.Cells(crtRow - i, 2).text
            
            If h.sMark.test(crtTgt) Then
                tgtSectDist = i
                tgtFoundSect = True
                
                If srcFoundSect Then
                    Exit For
                    
                End If
            End If
        End If
    
    Next i
    
    If srcSectDist = tgtSectDist Then
        adjust_to_previous_mark = crtRow - srcSectDist
    
    ElseIf srcSectDist > tgtSectDist Then
        ws.Range(Cells(crtRow - srcSectDist, 1), Cells(crtRow - tgtSectDist - 1, 1)).Select
        Selection.Insert Shift:=xlDown
        adjust_to_previous_mark = crtRow - srcSectDist + 1
    
    ElseIf srcSectDist < tgtSectDist Then
        ws.Range(Cells(crtRow - tgtSectDist, 2), Cells(crtRow - srcSectDist - 1, 2)).Select
        Selection.Insert Shift:=xlDown
        adjust_to_previous_mark = crtRow - tgtSectDist + 1
    
    End If

End Function

' 直近の下側の区切りが一致しているかのテスト
' 原文側、訳文側それぞれセルの距離を測り、差異のある分だけ調整する
'調整が終わったら調整後の行数を返す
Private Function adjust_to_next_mark(ByVal crtRow As Integer) As Integer
    
    Dim i As Integer
    Dim crtSrc As String, crtTgt As String
    Dim srcFoundSect As Boolean, tgtFoundSect As Boolean
    Dim srcSectDist As Integer, tgtSectDist As Integer
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    For i = 0 To h.modify_limit
                
        If srcFoundSect = False Then
            crtSrc = ws.Cells(crtRow + i, 1).text
            If h.sMark.test(crtSrc) Or crtSrc = FILE_END_MARK Then
                srcSectDist = i
                srcFoundSect = True
                
                If tgtFoundSect Then
                    Exit For
                    
                End If
            End If
        End If
        
        If tgtFoundSect = False Then
            crtTgt = ws.Cells(crtRow + i, 2).text
            If h.sMark.test(crtTgt) Or crtTgt = FILE_END_MARK Then
                tgtSectDist = i
                tgtFoundSect = True
                
                If srcFoundSect Then
                    Exit For
                    
                End If
            End If
        End If
    
    Next i
    
    
    If srcSectDist = tgtSectDist Then
        If crtSrc = FILE_END_MARK Or crtTgt = FILE_END_MARK Then
            adjust_to_next_mark = crtRow + h.modify_limit + 1
            
        Else
            adjust_to_next_mark = crtRow + srcSectDist
        
        End If
    
    ElseIf srcSectDist > tgtSectDist Then
        ws.Range(Cells(crtRow + srcSectDist - 1, 2), Cells(crtRow + tgtSectDist, 2)).Select
        Selection.Insert Shift:=xlDown
        
        If crtSrc = FILE_END_MARK Or crtTgt = FILE_END_MARK Then
            adjust_to_next_mark = crtRow + h.modify_limit + 1
        
        Else
            adjust_to_next_mark = crtRow + srcSectDist
        
        End If
    
    ElseIf srcSectDist < tgtSectDist Then
        ws.Range(Cells(crtRow + tgtSectDist - 1, 1), Cells(crtRow + srcSectDist, 1)).Select
        Selection.Insert Shift:=xlDown
        
        If crtSrc = FILE_END_MARK Or crtTgt = FILE_END_MARK Then
            adjust_to_next_mark = crtRow + h.modify_limit + 1
            
        Else
            adjust_to_next_mark = crtRow + tgtSectDist
            
        End If
    
    End If

End Function

'調整範囲内で原文・訳文ともに空白の行を削除する
Public Sub blank_row_delete(ByVal crtRow As Integer)

    Dim srcStr As String, tgtStr As String
    Dim i As Integer, j As Integer
    Dim startRow As Integer
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    '行の削除により文字列を取得する行がずれるため
    '常にインクリメントする j と
    '削除しなかったときのみインクリメントをする i を用意
    i = crtRow
    j = 0
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    srcStr = ws.Cells(i, 1).text
    tgtStr = ws.Cells(i, 2).text
    
    For j = 0 To h.modify_limit
        If srcStr = "" And tgtStr = "" Then
            ws.Rows(i).EntireRow.Delete
        
        ' EOF記号にあたったら削除を終了する
        ElseIf srcStr = FILE_END_MARK Or tgtStr = FILE_END_MARK Then
            'j = h.modify_limit
            Exit For
        
        Else
            i = i + 1
        
        End If
        
        srcStr = ws.Cells(i, 1).text
        tgtStr = ws.Cells(i, 2).text
    
    Next j

End Sub

Private Sub blank_row_delete_to_next(ByVal crtRow As Integer)
    
    Dim srcStr As String, tgtStr As String
    Dim i As Integer, j As Integer
    
    '行の削除により文字列を取得する行がずれるため
    '常にインクリメントする j と
    '削除しなかったときのみインクリメントをする i を用意
    i = crtRow
    j = 0
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    For j = 0 To h.modify_limit
        srcStr = ws.Cells(i, 1).text
        ' 区切り記号にあたったら削除を終了する
        If h.sMark.test(srcStr) Then
            Exit For
        
        ElseIf h.fMark.test(srcStr) Then
            Exit For
            
        ElseIf srcStr = FILE_END_MARK Then
            Exit For
        
        End If
        
        tgtStr = ws.Cells(i, 2).text
        
        If srcStr = "" And tgtStr = "" Then
            ws.Rows(i).EntireRow.Delete
        
        Else
            i = i + 1
        
        End If
            
    Next j

End Sub

'比較を実行する

Public Sub compare_vals_by_condition(ByVal cType As compdition)

    Dim crtRow As Integer, crtCol As Integer
    Dim i As Integer
    Dim srcStr As String, tgtStr As String
    Dim compare_start As Integer, compare_end As Integer, compare_finish As Integer
    
    Dim cMark As Variant
    
    Call standby
        
    If ActiveSheet.name <> "ALIGN" Then
        MsgBox ("ALIGN シートで実行してください")
        Exit Sub
        
    End If
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    '比較のために一時的に使用するC～E列をクリア
    ws.Range("C:E").Clear
    
    '比較終了後に戻ってくるため、現在の行・列数を取得
    crtRow = ActiveCell.Row
    crtCol = ActiveCell.Column
    
    'cType が FULL の場合は2行目から比較を開始する
    If cType = compdition.Full Then
        compare_start = 2
        compare_finish = h.end_row
    
    'cType が FULL 以外の場合はまず上方向に条件に合致するセルを探索
    Else
        Set cMark = CreateObject("VBScript.RegExp")
        If cType = compdition.Section Then
            cMark.pattern = SECTION_PATTERN
            
        Else
            cMark.pattern = FILE_PATTERN
            
        End If
        
        For i = crtRow To 1 Step -1
            srcStr = ws.Cells(i, 1).text
            
            If cMark.test(srcStr) Then
                If cMark.test(ws.Cells(i, 2).text) Then
                    compare_start = i
                    Exit For
                
                '開始行が区切り記号で原文・訳文が一致していなかった場合は処理を中断
                Else
                    MsgBox i & "行目の区切り記号が合致していません"
                    
                End If
                
            End If
        
        Next i
        
        compare_finish = compare_start + h.compare_limit
        
    '比較開始行の設定ここまで
    End If
    
    '比較開始行から比較範囲内で比較を実行する
    For i = compare_start + 1 To compare_finish
        
        srcStr = ws.Cells(i, 1).text
        tgtStr = ws.Cells(i, 2).text
        
        '区切り記号までで中断する場合の処理
        If compare_end_check(cType, srcStr, tgtStr) Then
            Exit For
        
        '原文と訳文の内容が同じかどうかを判定
        ElseIf srcStr = tgtStr Then
            
            '原文と訳文が同一で、区切り記号でもない場合、当該行のC列に SAME と記入する
            If h.sMark.test(srcStr) = False And h.fMark.test(srcStr) = False Then
                If srcStr <> FILE_END_MARK Then
                    ws.Cells(i, 3).Value = "SAME"
                
                End If
            
            End If
            
        End If
        
        '原文・訳文セットで同じ内容が重複して出現していないかを確認するため、
        '原文と訳文の文字列をつなげたものをD列に一時的に記入する
        ws.Cells(i, 4).Value = srcStr & tgtStr
    
    Next i
    
    '後の削除操作で使用するため、比較を終了した行を記録しておく
    compare_end = i
    
    Call h.set_compared_params(compare_start, compare_end)
    
    '重複の確認　ここから
    
    '原文と訳文の文字列をつなげたものに対し、同じものが複数回出現していないかを確認するため、
    'E列をCOUNTIF 関数で埋める
    ws.Cells(compare_start, 5).Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(R2C[-1]:RC[-1],RC[-1])"
    ws.Range(Cells(compare_start, 5), Cells(compare_end - 2, 5)).Select
    Selection.FillDown
    
    'COUNTIFの結果を値にして固定
    ws.Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '原文と訳文の文字列をつないでいたD列を削除しておく
    ws.Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    
    '区切り記号は当然に重複して出現するため、カウント結果を削除しておく
    ws.Columns("A:A").AutoFilter Field:=1, _
       Criteria1:="=_@λ_*_λ@_", Operator:=xlAnd
    ws.Columns("D:D").Clear
    
    ws.Columns("A:A").AutoFilter Field:=1, _
       Criteria1:="=_@@_*", Operator:=xlAnd
    ws.Columns("D:D").Clear
    
    ws.ShowAllData
    
    'D列に対し、セルの内容が 1 であるものを完全一致で検索し、空白で置換（削除）
    Columns("D:D").Select
    Selection.Replace what:="1", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    '完全一致検索のチェックを外しておく
    Selection.Replace what:="", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Application.ScreenUpdating = True
    
    'ヘッダーが消えている可能性があるので再度記入する
    Call write_header
    
    '元のセルに戻る
    ws.Cells(crtRow, crtCol).Select
    
End Sub

Private Function compare_end_check(ByVal cType As compdition, ByVal srcStr As String, ByVal tgtStr As String) As Boolean

    If cType = compdition.Full Then
        compare_end_check = False
        
    ElseIf cType = compdition.file Then
        If srcStr = FILE_END_MARK Then
            compare_end_check = True
        
        End If
        
    Else
        If h.sMark.test(srcStr) Or h.sMark.test(tgtStr) Then
            compare_end_check = True
            
        End If
        
    End If

End Function


Public Sub delete_by_condition(ByVal condition As deldition)

    Dim i As Integer, answer As Integer
    Dim question As String
    
    Call standby
    
    If ActiveSheet.name <> "ALIGN" Then
        MsgBox ("ALIGN シートで実行してください")
        Exit Sub
        
    End If
    
    '削除を行ってもいいかの確認
    '削除動作は一回の比較につき、一回だけ有効
    If h.able_to_del = False Then
        MsgBox ("削除前に比較を実行してください")
        Exit Sub
        
    Else
        Select Case condition
    
        Case deldition.same
            question = "原文＝訳文の"
            
        Case deldition.dupli
            question = "重複の"
            
        Case deldition.SAME_DUPLI
            question = "原文＝訳文 かつ 重複の"
        
        End Select
        
        answer = MsgBox(question & "削除を実行します。よろしいですか？", vbYesNo)
        If answer <> 6 Then
            Exit Sub
            
        Else
            h.able_to_del = False
            
        End If
    
    End If
    
    If makebu Then
       Call backup_align_sheet
        
    End If
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    Application.ScreenUpdating = False
    
    Select Case condition
    
        Case deldition.same
            For i = h.compare_start To h.compare_end
            
                If ws.Cells(i, 3) = "SAME" Then
                    ws.Rows(i).EntireRow.Delete
                    i = i - 1
                    
                End If
    
            Next i
            
        Case deldition.dupli
            For i = h.compare_start To h.compare_end
            
                If ws.Cells(i, 4) <> "" And ws.Cells(i, 4) <> "0" Then
                    ws.Rows(i).EntireRow.Delete
                    i = i - 1
                    
                End If
    
            Next i
            
        Case deldition.SAME_DUPLI
            For i = h.compare_start To h.compare_end
            
                If ws.Cells(i, 3) = "SAME" And ws.Cells(i, 4) <> "" And ws.Cells(i, 4) <> "0" Then
                    ws.Rows(i).EntireRow.Delete
                    i = i - 1
                    
                End If
    
            Next i
        
    End Select

    Application.ScreenUpdating = True

End Sub

Public Sub call_finish()

    Dim answer As Integer
    Dim isAddName As Boolean
    Dim i As Integer
    Dim srcStr As String, tgtStr As String, fileName As String
    Dim shp As Shape

    answer = MsgBox("C列にファイル名を追加しますか", vbYesNoCancel)
    
    If answer = 2 Then
        MsgBox ("キャンセルされました")
        Exit Sub

    ElseIf answer = 6 Then
        isAddName = True
    
    Else
        isAddName = False
    
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Sheets
    
        If ws.name = "FINISH" Then
            ws.Delete
            Exit For
        
        End If
        
    Next ws
    
    Sheets("ALIGN").Select
    Sheets("ALIGN").Copy After:=Sheets(2)
    Sheets("ALIGN (2)").Select
    Sheets("ALIGN (2)").name = "FINISH"
    Sheets("bu").Visible = True
    Sheets("FINISH").Select
    Sheets("FINISH").Move Before:=Sheets(1)
    Sheets("FINISH").Select
    Sheets("bu").Visible = False
    
    Set ws = ThisWorkbook.Sheets("FINISH")
    
    ws.Range("C:E").Clear
    
    For Each shp In ws.Shapes
        shp.Delete
        
    Next shp
    
    i = 2
    
    srcStr = ws.Cells(i, 1).text
    tgtStr = ws.Cells(i, 2).text
    fileName = ""
    
    Do While srcStr <> "" Or tgtStr <> ""
    
        If h.fMark.test(srcStr) Then
            
            If h.fMark.test(tgtStr) Then
                
                If isAddName Then
                    fileName = Replace(srcStr, "_@@_ ", "")
                    
                End If
            
                ws.Rows(i).EntireRow.Delete
                
            Else
            
                MsgBox ("ファイル名が一致しません")
                ws.Cells(i, 1).Select
                Exit Sub
            
            End If
        
        ElseIf h.sMark.test(srcStr) Then
            
            If h.sMark.test(tgtStr) Then
                
                ws.Rows(i).EntireRow.Delete
                
            Else
            
                MsgBox ("セグメントのタイプが一致しません")
                ws.Cells(i, 1).Select
                Exit Sub
            
            End If
                
        ElseIf srcStr = FILE_END_MARK Then
            
            If tgtStr = FILE_END_MARK Then
                
                ws.Rows(i).EntireRow.Delete
                fileName = ""
            
            Else
            
                MsgBox ("ファイルの終わりが一致しません")
                ws.Cells(i, 1).Select
                Exit Sub
            
            End If
            
        Else
            
            ws.Cells(i, 3) = fileName
            i = i + 1
            
        End If
        
        srcStr = ws.Cells(i, 1).text
        tgtStr = ws.Cells(i, 2).text
    
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox ("仕上げ処理が完了しました")

End Sub

Public Sub backup_align_sheet()

    Dim crtRow As Integer, crtCol As Integer
    
    crtRow = ActiveCell.Row
    crtCol = ActiveCell.Column

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Sheets
    
        If ws.name = "bu" Then
            ws.Delete
            Exit For
        
        End If
        
    Next ws
    
    Sheets("ALIGN").Select
    Sheets("ALIGN").Copy After:=Sheets(2)
    Sheets("ALIGN (2)").Select
    Sheets("ALIGN (2)").name = "bu"
    
    With Sheets("bu")
        .Move Before:=Sheets(1)
        .Range("A1:E1").Interior.Color = 255
        .Cells(crtRow, crtCol).Select
        .Visible = False
    
    End With
    
    Sheets("ALIGN").Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Public Sub restore_align_sheet()
    
    Dim answer As Integer
    Dim hasBu As Boolean
    Dim crtRow As Integer, crtCol As Integer
    
    hasBu = False
    For Each ws In ThisWorkbook.Sheets
    
        If ws.name = "bu" Then
            hasBu = True
            Exit For
        
        End If
        
    Next ws
    
    If hasBu = False Then
    
        MsgBox ("バックアップが存在しません")
        Exit Sub
        
    End If
    
    answer = MsgBox("バックアップから復元します。現在のALIGNシートは削除されます。よろしいですか？", vbYesNo)
    
    If answer = 7 Then
        Exit Sub
        
    End If
    
    crtRow = ActiveCell.Row
    crtCol = ActiveCell.Column
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Sheets
    
        If ws.name = "ALIGN" Then
            ws.Delete
            Exit For
        
        End If
        
    Next ws
    
    Sheets("bu").Visible = True
    Sheets("bu").name = "ALIGN"
    Sheets("ALIGN").Select
    Sheets("ALIGN").Range("A1:E1").Interior.pattern = xlNone
    
    Call backup_align_sheet
    
    Sheets("ALIGN").Cells(crtRow, crtCol).Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Public Sub test()

    Call standby

    Debug.Print (CatovisH.inculedFiles.ListIndex)
    
End Sub

Public Sub reset_all()

    Dim ws As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call MAIN.start_preparation
    Call h.reset_params
    
    For Each ws In ThisWorkbook.Sheets
    
        If ws.name = "FINISH" Or ws.name = "bu" Then
            ws.Delete

        End If
        
    Next ws
    
    'Call set_files_inBar("")
    
    CatovisH.inculedFiles.Clear
    CatovisH.inculedFiles2.Clear

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Public Sub abstract()

    Dim response As String, response_() As String
    
    Dim tms() As String, tbs() As String
    Dim tm As Variant, tb As Variant
    Dim tm_() As String, tb_() As String
    
    Dim i As Integer

    Call standby
    
    response = req.inqure_abstract
    
    response_ = Split(response, "&")
    
    tms = Split(response_(0), ",")
    tbs = Split(response_(1), ",")
    
    CATOVIS_ABS.TMListBox.Clear
    
    i = 0
    For Each tm In tms
        If tm <> "" Then
            tm_ = Split(tm, ":")
            CATOVIS_ABS.TMListBox.AddItem ("")
            CATOVIS_ABS.TMListBox.list(i, 0) = tm_(0)
            CATOVIS_ABS.TMListBox.list(i, 1) = tm_(1)
            i = i + 1
            
        End If
    
    Next tm
    
    i = 0
    For Each tb In tbs
        If tb <> "" Then
            tb_ = Split(tb, ":")
            CATOVIS_ABS.TBListBox.AddItem ("")
            CATOVIS_ABS.TBListBox.list(i, 0) = tb_(0)
            CATOVIS_ABS.TBListBox.list(i, 1) = tb_(1)
            i = i + 1
            
        End If
    
    Next tb
    
    CATOVIS_ABS.Show
    
End Sub

Public Sub get_tmtbData(ByVal dbType As String, ByVal fid As Integer, ByVal listVal As String)

    Dim response As String, response_() As String
    Dim lines() As String, lineBuf() As String
    Dim line As Variant, data As Variant
    Dim i As Long
    
    If dbType <> "TMs" And dbType <> "TBs" And dbType <> "all" Then
        Exit Sub
        
    End If

    response = req.get_byFid("?dbType=" & dbType & "&fid=" & fid)
    lines = Split(response, "[;]")
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    With ws
        .Select
        .Range(Cells(2, 1), Cells(h.end_row, 1)).EntireRow.Delete
        
        With .Range("A:E")
            .Interior.pattern = xlNone
            .NumberFormatLocal = "@"
        End With
    
    End With
    
    ReDim rngdata(99999, 1)
    
    i = 1
    rngdata(0, 0) = "原文"
    rngdata(0, 1) = "訳文"
    
    For Each line In lines
        If line <> "" Then
            lineBuf = Split(line, "[->]")
            rngdata(i, 0) = lineBuf(0)
            rngdata(i, 1) = lineBuf(1)
        
        End If
        
        i = i + 1
    
    Next line
    
    ws.Range("A1").Resize(i, 2) = rngdata
    
    'ヘルパークラスのプロパティをセットする
    Call h.set_status(listVal, "", i)
    'ステータスのシートを更新
    Call h.write_status
    
End Sub
