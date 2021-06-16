VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cat_Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BASE As String
Public PORT As String
Public URL As String

Public srcFile As String
' Public tgtFile As String

Public res As String

Sub Class_Initialize()
    
    Call settings("http://localhost", "8686")
    
End Sub

Sub Class_Terminate()

End Sub

Public Sub settings(ByVal base_ As String, ByVal port_ As String)

    BASE = base_
    PORT = port_
    URL = BASE & ":" & PORT & "/"

End Sub

Private Function req_sender(ByVal method As String, ByVal endpoint As String, ByVal contype As String, ByVal val As Variant) As String

    Dim req As XMLHTTP60
    Dim ep() As String
    
    Set req = New XMLHTTP60
    With req
        .Open method, URL + endpoint
        .setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"
        .setRequestHeader "Content-Type", contype
        .setRequestHeader "Charset", "utf-8"
        .setRequestHeader "User-Agent", "VBA"
        .send (val)
    
        Do While req.readyState < 4
            DoEvents
        Loop
    End With
    
    req_sender = req.responseText

End Function

Private Function jsonify(ByVal st As String, ByVal tt As String, Optional ByVal from As String = "") As String

    ' Chr(34)はダブルクォーテーションのこと
    
    Dim fileName As String
    Dim src As String, tgt As String
    
    If from = "" Then
        fileName = ActiveWindow.Caption
    Else
        fileName = strip_invalid_chr(from)
        
    End If
    
    src = strip_invalid_chr(st)
    tgt = strip_invalid_chr(tt)
    
    jsonify = _
    "{" & _
        Chr(34) & "src" & Chr(34) & ": " & Chr(34) & src & Chr(34) & "," & _
        Chr(34) & "tgt" & Chr(34) & ": " & Chr(34) & tgt & Chr(34) & "," & _
        Chr(34) & "from" & Chr(34) & ": " & Chr(34) & fileName & Chr(34) & _
    "}"

End Function

Private Function strip_invalid_chr(ByVal text As String) As String

    Dim text_ As String
    
    text_ = text
    text_ = Replace(text_, vbCr, "")
    text_ = Replace(text_, vbLf, "")
    text_ = Replace(text_, vbCrLf, "")
    text_ = Replace(text_, vbFormFeed, "")
    text_ = Replace(text_, Chr(7), "")
    
    strip_invalid_chr = text_

End Function

Public Function inqure_abstract() As String

    res = req_sender("GET", "abstract", "text/plain", "")
    
    inqure_abstract = res

End Function

Public Function get_byFid(ByVal fidQue As String) As String

    Dim qs As String
    
    qs = "status" & fidQue

    get_byFid = req_sender("GET", qs, "text/plain", "")

End Function

