VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CatovisH 
   Caption         =   "CATOVIS Helper"
   ClientHeight    =   5115
   ClientLeft      =   24120
   ClientTop       =   6465
   ClientWidth     =   3150
   OleObjectBlob   =   "CatovisH.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "CatovisH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub userform_initialize()

    'PASS

End Sub

Private Sub startBtn_Click()

    Call MAIN.open_align_file

End Sub

Private Sub finishBtn_Click()

    Call MAIN.call_finish

End Sub

Private Sub mergeBtn_Click()

    Call MAIN.lint_merge

End Sub

Private Sub splitBtn_Click()
    
    Call MAIN.lint_split

End Sub

Private Sub insertBtn_Click()

    Call MAIN.lint_insert

End Sub

Private Sub adjustBtn_Click()
    
    Call MAIN.lint_adjust

End Sub

Private Sub moveBtn_Click()

    Dim name As String
    
    name = inculedFiles.text
    
    If name <> "" Then
    
        Call MAIN.move_to_file(name)
        
    End If

End Sub

Private Sub moveBtn2_Click()

    Dim name As String
    
    name = inculedFiles2.text

    If name <> "" Then
        
        Call MAIN.move_to_file(name)
        
    End If

End Sub

Private Sub compSectBtn_Click()

    Call MAIN.compare_vals_by_condition("SECTION")

End Sub

Private Sub compFileBtn_Click()

    Call MAIN.compare_vals_by_condition("FILE")

End Sub

Private Sub compFullBtn_Click()

    Call MAIN.compare_vals_by_condition("FULL")

End Sub


Private Sub blankDelBtn_Click()

    Call MAIN.blank_row_delete(ActiveCell.Row)

End Sub

Private Sub dupliDelBtn_Click()

    Call MAIN.delete_by_condition("DUPLI")

End Sub

Private Sub sameDelBtn_Click()

    Call MAIN.delete_by_condition("SAME")

End Sub

Private Sub sameDupliDelBtn_Click()

    Call MAIN.delete_by_condition("SAME_DUPLI")

End Sub

Private Sub buBtn_Click()

    Dim answer As Integer
    
    answer = MsgBox("バックアップを作成しますか", vbYesNo)
    
    Debug.Print answer
    
    If answer = 6 Then
        Call MAIN.backup_align_sheet
    
    End If

End Sub

Private Sub restoreBtn_Click()

    Call MAIN.restore_align_sheet

End Sub

Private Sub setBuBtn_Click()

    Call MAIN.set_backup

End Sub
