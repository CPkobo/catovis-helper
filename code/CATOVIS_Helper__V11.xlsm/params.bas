Attribute VB_Name = "params"
Option Explicit

'���ʂ̒萔�E�ϐ����`����
'��؂蕶���i������j
'�t�B���^�Ŏg����悤�A���C���h�J�[�h������
Public Const WORD_FILE_MARK = "_@@_*.docx"
Public Const EXCEL_FILE_MARK = "_@@_*.xlsx"
Public Const PPT_FILE_MARK = "_@@_*.pptx"

'�t�@�C���I���}�[�N
Public Const FILE_END_MARK = "_@@_ EOF"

'�t�@�C���̋�؂蕶�������肷�邽�߂̐��K�\��
Public Const FILE_PATTERN = "^_@@_.+\.docx|_@@_.+\.xlsx|_@@_.+\.pptx$"

'�i���E�e�[�u���E�X���C�h�E�V�[�g�ȂǁA�������͈͂̋�؂蕶��
'�t�B���^�Ŏg����悤�A���C���h�J�[�h������
Public Const WORD_PARA_MARK = "_@��_ PARAGRAPH _��@_"
Public Const WORD_TBOX_MARK = "_@��_ TEXTBOX _��@_"
Public Const WORD_TABLE_MARK = "_@��_ TABLE _��@_"

Public Const EXCEL_SHEET_MARK = "_@��_ SHEET*_��@_"

Public Const PPT_SLIDE_MARK = "_@��_ SLIDE*_��@_"

'���͈͂̋�؂蕶�������肷�邽�߂̐��K�\��
Public Const SECTION_PATTERN = "^_@��_.+_��@_$"

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

