Option Explicit

' For debugging
Global Debugging As Boolean
Global FileNum As Integer

Sub WriteLine(line As String)
    If line <> Empty Then
        If Debugging Then
            Debug.Print line
        Else
            Print #FileNum, line
        End If
    End If
End Sub

Sub AlertError(msg As String)
    If Debugging Then
        Debug.Print "ERROR", msg
    Else
        MsgBox ("An error occurred:" & vbNewLine & msg)
    End If
End Sub

Sub AlertMessage(msg As String)
    If Debugging Then
        Debug.Print "MESSAGE", msg
    Else
        MsgBox (msg)
    End If
End Sub
