''' XML element functions
Function CreateElement(tag As String, value As String, Optional keepIfEmpty As Boolean) As String
    Dim val As String
    val = Trim(value)
    
    If keepIfEmpty Or val <> Empty Then
        CreateElement = "<" & tag & ">"
        CreateElement = CreateElement & ReplaceEntities(value)
        CreateElement = CreateElement & "</" & tag & ">"
    End If
End Function

Function WrapElement(parentTag As String, child As String, Optional keepIfEmpty As Boolean) As String
    Dim val As String
    val = Trim(child)
    
    If keepIfEmpty Or child <> Empty Then
        WrapElement = "<" & parentTag & ">" & vbNewLine
        WrapElement = WrapElement & child & vbNewLine
        WrapElement = WrapElement & "</" & parentTag & ">"
    End If
End Function

Function CreateParentElement(parentTag, children As Collection) As String
    CreateParentElement = "<" & parentTag & ">" & vbNewLine
    
    Dim child As Variant
    For Each child In children
        CreateParentElement = CreateParentElement & child
    Next
    
    CreateParentElement = CreateParentElement & "</" & parentTag & ">"
End Function

Function ReplaceEntities(value As String) As String
    If IsNull(value) Then
        ReplaceEntities = ""
    Else
        ReplaceEntities = value
        ReplaceEntities = Replace(ReplaceEntities, "&", "&amp;")
        ReplaceEntities = Replace(ReplaceEntities, """", "&quot;")
        ReplaceEntities = Replace(ReplaceEntities, "<", "&lt;")
        ReplaceEntities = Replace(ReplaceEntities, ">", "&gt;")
        ReplaceEntities = Replace(ReplaceEntities, "'", "&apos;")
    End If
End Function
