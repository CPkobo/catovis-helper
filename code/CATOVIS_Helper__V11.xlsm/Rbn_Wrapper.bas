Attribute VB_Name = "Rbn_Wrapper"
Option Explicit

'Callback for importFileBtn onAction
Sub btn_importFile(control As IRibbonControl)
    
    Call MAIN.open_align_file

End Sub

'Callback for finishBtn onAction
Sub btn_finish(control As IRibbonControl)

    Call MAIN.call_finish

End Sub

'Callback for importLSBtn onAction
Sub btn_importLS(control As IRibbonControl)

    Call MAIN.abstract

End Sub

'Callback for exportLSBtn onAction
Sub btn_exportLS(control As IRibbonControl)

    MsgBox "開発中です。今しばらくお待ちください"

End Sub

'Callback for lintMergeBtn onAction
Sub btn_merge(control As IRibbonControl)

    Call MAIN.lint_merge

End Sub

'Callback for lintSplitBtn onAction
Sub btn_split(control As IRibbonControl)

    Call MAIN.lint_split

End Sub

'Callback for lintInsertBtn onAction
Sub btn_insert(control As IRibbonControl)

    Call MAIN.lint_insert

End Sub

'Callback for lintAdjBtn onAction
Sub btn_adjust(control As IRibbonControl)

    Call MAIN.lint_adjust

End Sub

'Callback for fileMvBtn onAction
Sub btn_move(control As IRibbonControl)

    MsgBox "開発中です。今しばらくお待ちください"

'    Dim cb As CommandBar
'    Dim ctrl As Object
'    Dim name As String
'
'    For Each cb In CommandBars
'
'        If cb.name = "CATOVIS" Then
'            For Each ctrl In cb.Controls
'                If ctrl.Caption = "FILES" Then
'                    name = ctrl.text
'                    Exit For
'
'                End If
'
'            Next ctrl
'
'            Exit For
'        End If
'
'    Next cb
'
'    If name <> "" Then
'        Call MAIN.move_to_file(name)
'    End If
    
End Sub

'Callback for cprSepBtn onAction
Sub btn_comp_sep(control As IRibbonControl)

    Call MAIN.compare_vals_by_condition(Section)

End Sub

'Callback for cprFileBtn onAction
Sub btn_comp_file(control As IRibbonControl)

    Call MAIN.compare_vals_by_condition(file)

End Sub

'Callback for cprWhlBtn onAction
Sub btn_comp_whole(control As IRibbonControl)

    Call MAIN.compare_vals_by_condition(Full)

End Sub

'Callback for delBlkBtn onAction
Sub btn_del_blank(control As IRibbonControl)

    Call MAIN.blank_row_delete(ActiveCell.Row)
    
End Sub

'Callback for delDplBtn onAction
Sub btn_del_dupli(control As IRibbonControl)

    Call MAIN.delete_by_condition(dupli)

End Sub

'Callback for delSameBtn onAction
Sub btn_del_same(control As IRibbonControl)

    
    Call MAIN.delete_by_condition(same)

End Sub

'Callback for delDplSameBtn onAction
Sub btn_del_dupli_same(control As IRibbonControl)

    
    Call MAIN.delete_by_condition(SAME_DUPLI)

End Sub

'Callback for backupBtn onAction
Sub btn_backup(control As IRibbonControl)

    Call MAIN.backup_align_sheet

End Sub

'Callback for restoreBtn onAction
Sub btn_restore(control As IRibbonControl)

    Call MAIN.restore_align_sheet

End Sub

'Callback for dispFormBtn onAction
Sub btn_showForm(control As IRibbonControl)

    Call MAIN.call_helper

End Sub


