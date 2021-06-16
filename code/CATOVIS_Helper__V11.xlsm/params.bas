Attribute VB_Name = "params"
Option Explicit

'共通の定数・変数を定義する
'区切り文字（文字列）
'フィルタで使えるよう、ワイルドカード文字列
Public Const WORD_FILE_MARK = "_@@_*.docx"
Public Const EXCEL_FILE_MARK = "_@@_*.xlsx"
Public Const PPT_FILE_MARK = "_@@_*.pptx"

'ファイル終了マーク
Public Const FILE_END_MARK = "_@@_ EOF"

'ファイルの区切り文字か判定するための正規表現
Public Const FILE_PATTERN = "^_@@_.+\.docx|_@@_.+\.xlsx|_@@_.+\.pptx$"

'段落・テーブル・スライド・シートなど、小さい範囲の区切り文字
'フィルタで使えるよう、ワイルドカード文字列
Public Const WORD_PARA_MARK = "_@λ_ PARAGRAPH _λ@_"
Public Const WORD_TBOX_MARK = "_@λ_ TEXTBOX _λ@_"
Public Const WORD_TABLE_MARK = "_@λ_ TABLE _λ@_"

Public Const EXCEL_SHEET_MARK = "_@λ_ SHEET*_λ@_"

Public Const PPT_SLIDE_MARK = "_@λ_ SLIDE*_λ@_"

'小範囲の区切り文字か判定するための正規表現
Public Const SECTION_PATTERN = "^_@λ_.+_λ@_$"

Public Enum compdition

    Full
    file
    Section

End Enum

Public Enum deldition

    same
    dupli
    SAME_DUPLI
    
End Enum

