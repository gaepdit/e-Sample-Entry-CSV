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

''' Complex data types
Function SpecializedMeasurement(tag As String, value As Variant, Optional typeCode As String) As String
    ' "value" must be numeric
    If value = Empty Then Exit Function
    
    Dim children As New Collection
    children.Add CreateElement("EN:MeasurementValue", CStr(value))
    children.Add CreateElement("EN:MeasurementSignificantDigit", GetSigFigs(value))
    If typeCode <> Empty Then
        children.Add CreateElement("EN:SpecializedMeasurementTypeCode", typeCode)
    End If
    
    SpecializedMeasurement = CreateParentElement(tag, children)
End Function

Function UnitMeasurement(tag As String, value As Integer, units As String) As String
    Dim children As New Collection
    children.Add CreateElement("EN:MeasurementValue", CStr(value))
    children.Add CreateElement("EN:MeasurementUnit", units)
    
    UnitMeasurement = CreateParentElement(tag, children)
End Function

Function GetSigFigs(value As Variant) As Integer
    Dim val As String
    val = CStr(CDec(value))
    
    GetSigFigs = Len(val) - InStr(val, ".")
End Function

''' Utilities
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