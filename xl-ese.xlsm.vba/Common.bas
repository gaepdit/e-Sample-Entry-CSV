Option Explicit

' For debugging
Global production As Boolean

Sub WriteLine(line As String)
    If production Then
        Print #1, line
    Else
        Debug.Print line
    End If
End Sub

Sub AlertError(msg As String)
    If production Then
        MsgBox ("An error occurred:" & vbNewLine & msg)
    Else
        Debug.Print "ERROR", msg
    End If
End Sub

Sub AlertMessage(msg As String)
    If production Then
        MsgBox (msg)
    Else
        Debug.Print "MESSAGE", msg
    End If
End Sub

Function GetSigFigs(value As Variant) As Integer
    Dim val As String
    val = CStr(CDec(value))
    
    GetSigFigs = Len(val) - InStr(val, ".")
End Function