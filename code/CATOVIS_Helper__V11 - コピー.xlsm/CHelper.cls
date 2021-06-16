VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fMark As Variant
Public sMark As Variant
Public eofMark As Variant

Public modify_limit As Integer
Public compare_limit As Integer

Public end_row As Integer

Public compare_start As Integer
Public compare_end As Integer
Public able_to_del As Boolean

Private align_file As String
Private including_files() As String
Private including_files_num As Integer

Private Sub Class_Initialize()

    Dim align As String, files As String
    Dim end_ As Integer

    Call create_regexpObj
    Call set_limits
    
    end_ = ThisWorkbook.Sheets("ALIGN").Range("A1").SpecialCells(xlLastCell).Row
    
    With ThisWorkbook.Sheets("STATUS")
    
        .Range("B3") = end_
    
        If .Range("B2") <> "" Then
            align = .Range("B2")
            files = .Range("B4")
            
        Else
            align = ""
            files = ""
            
        End If
        
        Call set_status(align, files, end_)
        
    End With

End Sub

Public Sub reset_params()

    end_row = 1
    compare_start = 0
    compare_end = 0
    align_file = ""
    ReDim including_files(0)
    including_files_num = 0

End Sub

Private Sub create_regexpObj()

    Set fMark = CreateObject("VBScript.RegExp")
    fMark.pattern = params.FILE_PATTERN
    
    Set sMark = CreateObject("VBScript.RegExp")
    sMark.pattern = params.SECTION_PATTERN
    
    Set eofMark = CreateObject("VBScript.RegExp")
    eofMark.pattern = params.FILE_END_MARK

End Sub

Public Sub set_status(ByVal align As String, ByVal files As String, ByVal end_ As Integer)

    align_file = align
    including_files = Split(files, vbCrLf)
    including_files_num = UBound(including_files) + 1
    end_row = end_

End Sub

Public Sub write_status()

    Dim i As Integer

    With ThisWorkbook.Sheets("STATUS")
        .Range("B2") = align_file
        .Range("B3") = end_row
        .Range("B4") = Join(including_files, vbCrLf)
        .Range("B5") = including_files_num
        .Range("B2:B5").HorizontalAlignment = xlLeft
    End With

End Sub

Public Sub set_limits()

    modify_limit = ThisWorkbook.Sheets("PREFERENCE").Range("E4").Value
    compare_limit = ThisWorkbook.Sheets("PREFERENCE").Range("E5").Value

End Sub

Public Sub set_compared_params(ByVal start_ As Integer, ByVal end_ As Integer)
    
    compare_start = start_
    compare_end = end_
    able_to_del = True

End Sub

Public Sub compared_delete()

    able_to_del = False

End Sub

Property Get align_file_name() As String
    
    align_file_name = align_file

End Property

Property Get files_list() As Variant
        
    files_list = including_files

End Property



