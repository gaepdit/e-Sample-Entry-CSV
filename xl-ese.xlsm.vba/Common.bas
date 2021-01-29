Option Explicit

' For debugging
Global Production As Boolean
Global FileNum As Integer

Sub WriteLine(line As String)
    If line <> Empty Then
        If Production Then
            Print #FileNum, line
        Else
            Debug.Print line
        End If
    End If
End Sub

Sub AlertError(msg As String)
    If Production Then
        MsgBox ("An error occurred:" & vbNewLine & msg)
    Else
        Debug.Print "ERROR", msg
    End If
End Sub

Sub AlertMessage(msg As String)
    If Production Then
        MsgBox (msg)
    Else
        Debug.Print "MESSAGE", msg
    End If
End Sub
