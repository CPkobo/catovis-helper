Attribute VB_Name = "MAIN"
Option Explicit


Dim ws As Worksheet
Dim ready As Boolean
Dim makebu As Boolean
Dim h As CHelper
Dim req As Cat_Request

'�t�H�[�����Ăяo�����ۂɁA���[�N�V�[�g�Ƀt�H�[�J�X���c��悤�ɂ��邽�߂̃��C�u����
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

'CHelper�N���X���Ȃ���Ώ���������
Public Sub standby()

    If ready = False Then

        Set h = New CHelper
        Set req = New Cat_Request
        ready = True
        
    End If

End Sub

'�o�b�N�A�b�v�̐ݒ��ύX����
'��TODO �A�h�C���^�u������ύX�ł���悤�ɂ���
Public Sub set_backup()

    If ThisWorkbook.Sheets("PREFERENCE").Range("E6").Value = 1 Then
        makebu = True
    
    Else
        makebu = False
        
    End If

End Sub

'�}�N���̃V���[�g�J�b�g��ݒ肷��
Public Sub apply_shortcut()

    Set ws = ThisWorkbook.Sheets("PREFERENCE")

    Application.MacroOptions macro:="lint_merge", ShortcutKey:=ws.Range("C2").text
    Application.MacroOptions macro:="lint_split", ShortcutKey:=ws.Range("C3").text
    Application.MacroOptions macro:="lint_insert", ShortcutKey:=ws.Range("C4").text
    Application.MacroOptions macro:="lint_adjust", ShortcutKey:=ws.Range("C5").text

End Sub

'�w���p�[�t�H�[�����Ăяo��
Public Sub call_helper()
    
    If CatovisH.inculedFiles.ListCount = 0 Then
        Call set_files_inForm
    End If

    CatovisH.Show

End Sub

'�w���p�[�t�H�[���̃t�@�C�����X�g���X�V
Private Sub set_files_inForm()

    Dim list As Variant
    Dim i As Integer

    '�w���p�[�N���X����i�[���Ă���t�@�C�����X�g���擾
    list = h.files_list
    
    For i = 0 To UBound(list)
    
        CatovisH.inculedFiles.AddItem (list(i))
        CatovisH.inculedFiles2.AddItem (list(i))
    
    Next i

End Sub

'ALIGN�V�[�g�̈�s�ڂɃw�b�_�[����������
Public Sub write_header()

    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    ws.Range("A1").Value = "����"
    ws.Range("B1").Value = "��"
    ws.Range("C1").Value = "����"
    ws.Range("D1").Value = "�d��"

End Sub

'CATOVIS�̃A���C���t�@�C����ǂݍ���
'�t�H�[����START�{�^�����A�A�h�C���^�u�́u�J���v������s���邱�Ƃ�z��
Public Sub open_align_file()

    Call standby

    Dim strFiles As String, files As String
    Dim i As Integer

    '�t�@�C���̓ǂݍ��݃_�C�A���O�@��������
    With Application.FileDialog(msoFileDialogFilePicker)
        '�t�@�C���̕����I�����\�ɂ���
        .AllowMultiSelect = False
        '�t�@�C���t�B���^�̃N���A
        .Filters.Clear
        '�t�@�C���t�B���^�̒ǉ�
        .Filters.Add "���̑�", "*.tsv"
        '�����\���t�H���_�̐ݒ�
        .InitialFileName = ThisWorkbook.Path

        If .Show = -1 Then  '�t�@�C���_�C�A���O�\��
            ' [ OK ] �{�^���������ꂽ�ꍇ
            For i = 1 To .SelectedItems.Count
                strFiles = .SelectedItems(i)
            Next i

        Else
            ' [ �L�����Z�� ] �{�^���������ꂽ�ꍇ
            MsgBox "�t�@�C���I�����L�����Z������܂����B", vbExclamation
        End If
        
    End With
    
    '�t�@�C���̓ǂݍ��݃_�C�A���O�@�����܂�

    '�ǂݍ��񂾕������W�J����
    If strFiles <> "" Then

        Application.ScreenUpdating = False

        '�����̃f�[�^�������āA�����𕶎���ɂ���
        Call start_preparation
        
        '���ۂ�tsv��Excel�ϊ�����
        Call readin_TSV(strFiles)
    
        '�ǂݍ��񂾃f�[�^��l�̖ڂɌ��₷������
        Call decorate_friendly
        
        '�V�[�g�𕡐����Ă���
        Call backup_align_sheet
        
        '�t�H�[���̃t�@�C�����X�g���X�V����
        Call set_files_inForm
        
        Application.ScreenUpdating = True
        
        MsgBox ("�ǂݍ��݂��������܂���")
        
    End If
    
    '�t�H�[�J�X�����[�N�V�[�g�ɖ߂�
    SetFocus Application.hwnd

End Sub

'�t�@�C���̏����ݒ������
'��̓I�ɂ͈ȉ��̑���
'�t�@�C�����X�g�̍폜�A�f�[�^�̍폜�A�w�b�_�[�̒ǋL�A�t�B���^�[�̃Z�b�g�A�����̃Z�b�g
'�V�K�t�@�C���̓ǂݍ��ݎ��Ɏ����Ŏ��s�����
'�A�b�v���[�h�O�Ɏ蓮�Ŏ��s���邱�Ƃ���
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
        
        '�w�b�_�[������������
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

'�w�肳�ꂽTSV�t�@�C�����當����𒊏o���ăZ���ɑ������
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
    rngdata(0, 0) = "����"
    rngdata(0, 1) = "��"
    
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
                '�ǂݍ��񂾍s���t�@�C�����̏ꍇ�A�t�@�C�����X�g�Ƃ��ĕʕۊǂ���
                innerFiles = innerFiles + vbCrLf + Replace(rngdata(i, 0), "_@@_ ", "")
            
            End If
        
            i = i + 1
        
        End If
    
    Next line
    
    '���s�L�����擪�ɓ����Ă��܂��̂ō폜����
    innerFiles = Replace(innerFiles, vbCrLf, "", 1, 1)
    
    '�w���p�[�N���X�̃v���p�e�B���Z�b�g����
    Call h.set_status(target, innerFiles, i)
    '�X�e�[�^�X�̃V�[�g���X�V
    Call h.write_status
    '�A�h�C���^�u�̃t�@�C�����X�g���X�V
    'Call set_files_inBar(innerFiles)
    
    ws.Range("A1").Resize(i, 2) = rngdata
    
End Sub

'��؂蕶���ɐF�����Ă���
Private Sub decorate_friendly()

    Dim a As String
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    ws.Activate
 
 
    ' �t�@�C�����Ƃ�����i���ʃ��x���̋�؂�L���ɐF��t����
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
    If WorksheetFunction.CountIf(ws.Range("A:A"), "_@��_ SHEET1 _��@_") > 0 Then
        Call filter_and_color(EXCEL_FILE_MARK, 50)
        Call filter_and_color(EXCEL_SHEET_MARK, 43)
    End If
    
    'PPT
    If WorksheetFunction.CountIf(ws.Range("A:A"), "_@��_ SLIDE1 _��@_") > 0 Then
        Call filter_and_color(PPT_FILE_MARK, 46)
        Call filter_and_color(PPT_SLIDE_MARK, 45)
    End If
    
    
    ' EOF
    Call filter_and_color(FILE_END_MARK, 48)
    
    '�Ō�Ƀt�B���^����
    ws.ShowAllData
    ws.Range("A2").Select

End Sub

'�t�B���^�[�̏����ƐF���󂯎���Ĕ��f���邽�߂̓����v���V�[�W��
Private Sub filter_and_color(ByVal cr As String, ByVal cx As Long)

    Columns("A:A").Select
    Selection.AutoFilter Field:=1, _
        Criteria1:="=" & cr
    
    ws.Range(Cells(2, 1), Cells(h.end_row, 2)).Select
        
    Selection.Interior.ColorIndex = cx

End Sub

'�t�@�C�������󂯎���ăt�H�[�J�X���ړ�����T�u�v���V�[�W��
'�����@�\�𗘗p���Ă���
Public Sub move_to_file(ByVal name As String)

    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    Dim srcCol As Range
    Dim gotRow As Integer
    Dim nameWithMark As String
    
    Application.ScreenUpdating = False
    
    nameWithMark = "_@@_ " + name

    Set srcCol = ws.Columns("A:A")
    
    '�O�̂��߁u�Z���̊��S��v�v���g�p����
    If srcCol.Find(what:=nameWithMark, LookAt:=xlWhole) Is Nothing Then
        Exit Sub
    
    Else
        gotRow = srcCol.Find(what:=nameWithMark, LookAt:=xlWhole).Row
    
    End If
    
    '�Z���̊��S��v���������Ă���
    srcCol.Find(what:="", LookAt:=xlPart).Activate
    
    Application.ScreenUpdating = True
    
    ws.Cells(gotRow, 1).Select

End Sub

'�Z���̌�������
Public Sub lint_merge()
Attribute lint_merge.VB_ProcData.VB_Invoke_Func = "M\n14"

    Call standby

    Dim strInRange As String, strMerged As String
    Dim colsNum As Integer, rowsNum As Integer, i As Integer, j As Integer
    Dim crtRng As Range
    
    Set crtRng = Selection
    
    If crtRng.Rows.Count = 1 Then
        MsgBox ("2�s�ȏ�I�����Ă�������")
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
            
            '������^�łȂ���Ε�����^�ɕϊ����Ēǉ����
            If VarType(strInRange) <> 8 Then
                strMerged = strMerged + Str(strInRange)
            
            Else
                '��؂蕶���͌�������Ȃ��悤�ɂ���
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
        
        '��������
        Selection(1, 1).Value = strMerged
        'Selection.Resize(rowsNum - 1, 1).Select
        If j > 2 Then
            Selection.Resize(j - 2, 1).Select
            Selection.Offset(1, 0).Select
            Selection.Delete Shift:=xlShiftUp
        End If
        
    Next i

End Sub

'�I��͈͓��̃e�L�X�g���f���~�^�Ɋ�Â��ĕ���
'������̃Z�O�����g���Ɋ�Â��ĕ����̃Z���ɑ��
'#TODO �f���~�^�̕����Ή�
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
    
    '�����֎~���̃`�F�b�N
    Set crtRng = Selection
    
    '������1��݂̂ɑ΂��ėL��
    If crtRng.Columns.Count > 1 Then
        MsgBox "���1�񂾂���I�����Ă�������"
        End
        
    End If
    
    '�I��͈͂̍s���𐔂���
    rngNum = crtRng.Count
    
    '������20�s�܂�
    If rngNum > 1000 Then
        MsgBox "�s�̑I����20�s�܂łɂ��Ă�������"
        End
        
    End If
    
    '�f���~�^��ݒ肷��
    If crtRng.Column = 1 Then
        delimiter = ThisWorkbook.Sheets("PREFERENCE").Range("E2").text
    
    ElseIf crtRng.Column = 2 Then
        delimiter = ThisWorkbook.Sheets("PREFERENCE").Range("E3").text
        
    Else
        MsgBox "�����@�\�͌����� �܂��� �󕶗�ł̂ݗL���ł�"
        End
    
    End If
    
    '�֎~�ݒ�̊m�F�����܂�
    
    '�o�b�N�A�b�v���s��
    If makebu Then
        Call backup_align_sheet
    End If
    
    '��̕ϐ��Ɉ�U���ׂĂ̕�������i�[����
    allText = ""
    
    For j = 1 To rngNum
        
        singleText = crtRng(j).text
        
        ' ��؂�L���ɓ���������i�[�𒆒f����
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
    
    '�Ōオ�f���~�^���ǂ����̔�������Ă���
    isEndWithDelim = Right(allText, 1) = delimiter
    
    '�f���~�^�ŕ�������
    sentences = Split(allText, delimiter)
    sentenceNum = UBound(sentences)
    
    '�����̕K�v���Ȃ������ꍇ�͏������I��
    If sentenceNum = 0 Then
        Exit Sub
        
    End If
    
    'Split�ŏ������f���~�^��₤
    For i = 0 To sentenceNum - 1
        
        sentences(i) = sentences(i) + delimiter
        
    Next i
    
    '�Ō�̃Z�O�����g���f���~�^�ŏI����Ă����ꍇ�A��̕����񂪓����Ă���̂�
    'sentenceNum���f�N�������g
    
    If isEndWithDelim Then
        sentenceNum = sentenceNum - 1
        
    End If
    
    
    '������̃Z�O�����g�����I��͈͂�菭�Ȃ��ꍇ�A
    '�������ʂ��Z���ɋL�������̂��A�s�v�ȕ����폜����
    'If rngNum >= sentenceNum + 1 Then
    If j >= sentenceNum + 1 Then
        For i = 0 To sentenceNum
            Selection(i + 1).Value = sentences(i)
        Next i
        
        '�Ō�̃Z�O�����g�ȍ~�͓��e���폜
        'For i = sentenceNum + 1 To rngNum - 1
        For i = sentenceNum + 1 To j - 1
            Selection(i + 1).Clear
        Next i
        
    '������̃Z�O�����g�����I��͈͂�葽���ꍇ�A
    '�I��͈͈ȍ~�̓Z����}�����Ȃ��珑������
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

' �N���b�v�{�[�h�̒��g�������o���đ}������
Public Sub lint_insert()
Attribute lint_insert.VB_ProcData.VB_Invoke_Func = "I\n14"
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    Selection.Insert Shift:=xlDown
    ActiveSheet.Paste

End Sub

' ��؂�L���Ɋ�Â�����
'�����͌����ɖ󕶂����킹��`�ōs��
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
    
    ' ���݂̍s�����擾
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

' ���߂̏㑤�̋�؂肪��v���Ă��邩�̃e�X�g
' �������A�󕶑����ꂼ��Z���̋����𑪂�A���ق̂��镪������������
'�������I������璲����̍s����Ԃ�
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

' ���߂̉����̋�؂肪��v���Ă��邩�̃e�X�g
' �������A�󕶑����ꂼ��Z���̋����𑪂�A���ق̂��镪������������
'�������I������璲����̍s����Ԃ�
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

'�����͈͓��Ō����E�󕶂Ƃ��ɋ󔒂̍s���폜����
Public Sub blank_row_delete(ByVal crtRow As Integer)

    Dim srcStr As String, tgtStr As String
    Dim i As Integer, j As Integer
    Dim startRow As Integer
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    '�s�̍폜�ɂ�蕶������擾����s������邽��
    '��ɃC���N�������g���� j ��
    '�폜���Ȃ������Ƃ��̂݃C���N�������g������ i ��p��
    i = crtRow
    j = 0
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    srcStr = ws.Cells(i, 1).text
    tgtStr = ws.Cells(i, 2).text
    
    For j = 0 To h.modify_limit
        If srcStr = "" And tgtStr = "" Then
            ws.Rows(i).EntireRow.Delete
        
        ' EOF�L���ɂ���������폜���I������
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
    
    '�s�̍폜�ɂ�蕶������擾����s������邽��
    '��ɃC���N�������g���� j ��
    '�폜���Ȃ������Ƃ��̂݃C���N�������g������ i ��p��
    i = crtRow
    j = 0
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    For j = 0 To h.modify_limit
        srcStr = ws.Cells(i, 1).text
        ' ��؂�L���ɂ���������폜���I������
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

'��r�����s����

Public Sub compare_vals_by_condition(ByVal cType As compdition)

    Dim crtRow As Integer, crtCol As Integer
    Dim i As Integer
    Dim srcStr As String, tgtStr As String
    Dim compare_start As Integer, compare_end As Integer, compare_finish As Integer
    
    Dim cMark As Variant
    
    Call standby
        
    If ActiveSheet.name <> "ALIGN" Then
        MsgBox ("ALIGN �V�[�g�Ŏ��s���Ă�������")
        Exit Sub
        
    End If
    
    If makebu Then
        Call backup_align_sheet
    End If
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Sheets("ALIGN")
    
    '��r�̂��߂Ɉꎞ�I�Ɏg�p����C�`E����N���A
    ws.Range("C:E").Clear
    
    '��r�I����ɖ߂��Ă��邽�߁A���݂̍s�E�񐔂��擾
    crtRow = ActiveCell.Row
    crtCol = ActiveCell.Column
    
    'cType �� FULL �̏ꍇ��2�s�ڂ����r���J�n����
    If cType = compdition.Full Then
        compare_start = 2
        compare_finish = h.end_row
    
    'cType �� FULL �ȊO�̏ꍇ�͂܂�������ɏ����ɍ��v����Z����T��
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
                
                '�J�n�s����؂�L���Ō����E�󕶂���v���Ă��Ȃ������ꍇ�͏����𒆒f
                Else
                    MsgBox i & "�s�ڂ̋�؂�L�������v���Ă��܂���"
                    
                End If
                
            End If
        
        Next i
        
        compare_finish = compare_start + h.compare_limit
        
    '��r�J�n�s�̐ݒ肱���܂�
    End If
    
    '��r�J�n�s�����r�͈͓��Ŕ�r�����s����
    For i = compare_start + 1 To compare_finish
        
        srcStr = ws.Cells(i, 1).text
        tgtStr = ws.Cells(i, 2).text
        
        '��؂�L���܂łŒ��f����ꍇ�̏���
        If compare_end_check(cType, srcStr, tgtStr) Then
            Exit For
        
        '�����Ɩ󕶂̓��e���������ǂ����𔻒�
        ElseIf srcStr = tgtStr Then
            
            '�����Ɩ󕶂�����ŁA��؂�L���ł��Ȃ��ꍇ�A���Y�s��C��� SAME �ƋL������
            If h.sMark.test(srcStr) = False And h.fMark.test(srcStr) = False Then
                If srcStr <> FILE_END_MARK Then
                    ws.Cells(i, 3).Value = "SAME"
                
                End If
            
            End If
            
        End If
        
        '�����E�󕶃Z�b�g�œ������e���d�����ďo�����Ă��Ȃ������m�F���邽�߁A
        '�����Ɩ󕶂̕�������Ȃ������̂�D��Ɉꎞ�I�ɋL������
        ws.Cells(i, 4).Value = srcStr & tgtStr
    
    Next i
    
    '��̍폜����Ŏg�p���邽�߁A��r���I�������s���L�^���Ă���
    compare_end = i
    
    Call h.set_compared_params(compare_start, compare_end)
    
    '�d���̊m�F�@��������
    
    '�����Ɩ󕶂̕�������Ȃ������̂ɑ΂��A�������̂�������o�����Ă��Ȃ������m�F���邽�߁A
    'E���COUNTIF �֐��Ŗ��߂�
    ws.Cells(compare_start, 5).Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(R2C[-1]:RC[-1],RC[-1])"
    ws.Range(Cells(compare_start, 5), Cells(compare_end - 2, 5)).Select
    Selection.FillDown
    
    'COUNTIF�̌��ʂ�l�ɂ��ČŒ�
    ws.Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '�����Ɩ󕶂̕�������Ȃ��ł���D����폜���Ă���
    ws.Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    
    '��؂�L���͓��R�ɏd�����ďo�����邽�߁A�J�E���g���ʂ��폜���Ă���
    ws.Columns("A:A").AutoFilter Field:=1, _
       Criteria1:="=_@��_*_��@_", Operator:=xlAnd
    ws.Columns("D:D").Clear
    
    ws.Columns("A:A").AutoFilter Field:=1, _
       Criteria1:="=_@@_*", Operator:=xlAnd
    ws.Columns("D:D").Clear
    
    ws.ShowAllData
    
    'D��ɑ΂��A�Z���̓��e�� 1 �ł�����̂����S��v�Ō������A�󔒂Œu���i�폜�j
    Columns("D:D").Select
    Selection.Replace what:="1", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    '���S��v�����̃`�F�b�N���O���Ă���
    Selection.Replace what:="", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Application.ScreenUpdating = True
    
    '�w�b�_�[�������Ă���\��������̂ōēx�L������
    Call write_header
    
    '���̃Z���ɖ߂�
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
        MsgBox ("ALIGN �V�[�g�Ŏ��s���Ă�������")
        Exit Sub
        
    End If
    
    '�폜���s���Ă��������̊m�F
    '�폜����͈��̔�r�ɂ��A��񂾂��L��
    If h.able_to_del = False Then
        MsgBox ("�폜�O�ɔ�r�����s���Ă�������")
        Exit Sub
        
    Else
        Select Case condition
    
        Case deldition.same
            question = "�������󕶂�"
            
        Case deldition.dupli
            question = "�d����"
            
        Case deldition.SAME_DUPLI
            question = "�������� ���� �d����"
        
        End Select
        
        answer = MsgBox(question & "�폜�����s���܂��B��낵���ł����H", vbYesNo)
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

    answer = MsgBox("C��Ƀt�@�C������ǉ����܂���", vbYesNoCancel)
    
    If answer = 2 Then
        MsgBox ("�L�����Z������܂���")
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
            
                MsgBox ("�t�@�C��������v���܂���")
                ws.Cells(i, 1).Select
                Exit Sub
            
            End If
        
        ElseIf h.sMark.test(srcStr) Then
            
            If h.sMark.test(tgtStr) Then
                
                ws.Rows(i).EntireRow.Delete
                
            Else
            
                MsgBox ("�Z�O�����g�̃^�C�v����v���܂���")
                ws.Cells(i, 1).Select
                Exit Sub
            
            End If
                
        ElseIf srcStr = FILE_END_MARK Then
            
            If tgtStr = FILE_END_MARK Then
                
                ws.Rows(i).EntireRow.Delete
                fileName = ""
            
            Else
            
                MsgBox ("�t�@�C���̏I��肪��v���܂���")
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
    
    MsgBox ("�d�グ�������������܂���")

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
    
        MsgBox ("�o�b�N�A�b�v�����݂��܂���")
        Exit Sub
        
    End If
    
    answer = MsgBox("�o�b�N�A�b�v���畜�����܂��B���݂�ALIGN�V�[�g�͍폜����܂��B��낵���ł����H", vbYesNo)
    
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
    rngdata(0, 0) = "����"
    rngdata(0, 1) = "��"
    
    For Each line In lines
        If line <> "" Then
            lineBuf = Split(line, "[->]")
            rngdata(i, 0) = lineBuf(0)
            rngdata(i, 1) = lineBuf(1)
        
        End If
        
        i = i + 1
    
    Next line
    
    ws.Range("A1").Resize(i, 2) = rngdata
    
    '�w���p�[�N���X�̃v���p�e�B���Z�b�g����
    Call h.set_status(listVal, "", i)
    '�X�e�[�^�X�̃V�[�g���X�V
    Call h.write_status
    
End Sub
