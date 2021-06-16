VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CATOVIS_ABS 
   Caption         =   "UserForm1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940
   OleObjectBlob   =   "CATOVIS_ABS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CATOVIS_ABS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetTMBtn_Click()

    If TMListBox.ListIndex > -1 Then

        Call MAIN.get_tmtbData("TMs", TMListBox.ListIndex, TMListBox.Value & "(TM)")
    
    End If

End Sub

Private Sub GetTBBtn_Click()

    If TBListBox.ListIndex > -1 Then

        Call MAIN.get_tmtbData("TBs", TBListBox.ListIndex, TBListBox.Value & "(TB)")
    
    End If

End Sub

