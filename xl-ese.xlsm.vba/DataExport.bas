Option Explicit

Dim exportSamples As Boolean
Dim exportResults As Boolean

' For debugging
Dim production As Boolean

Sub GenerateXmlDocument()
    ' Temp while debugging
    exportSamples = True
    exportResults = False
    
    ' Set globals
    SetGlobals
    
    ' Ensure data exists
    If Not exportSamples And Not exportResults Then Exit Sub
    If Not exportSamples And TableIsEmpty(ResultsTable) Then Exit Sub
    If Not exportResults And TableIsEmpty(SamplesTable) Then Exit Sub
    
    ' For debugging
    If Not production Then
        Debug.Print
        Debug.Print Now
        Debug.Print "==="
    End If
    
    ' For error handling
    Dim closing As Boolean
    closing = False
    
    On Error GoTo ErrHandler:
    
' === Initial data validation (?)
'not done
    
' === Create file
'not done
    
    ' generate default file name/path
    ' request file name/path from user
    ' verify file does not exist or overwrite

    Dim fPath As String
    fPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & "export.xml"
    
    ' Open file for saving
    Open fPath For Output As #1
    
' === END Create file

' === Start document
    WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    WriteLine "<EN:eDWR xmlns:EN=""urn:us:net:exchangenetwork"" xmlns:SDWIS=""http://www.epa.gov/sdwis"" xmlns:ns2=""http://www.epa.gov/xml"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
    WriteLine "<EN:Submission EN:submissionFileCreatedDate=""" & Format(Now, "yyyy-mm-dd") & """>"
    WriteLine "<EN:LabReport>"
' === END Start document

' === Output LabIdentification
    WriteLine "<EN:LabIdentification>"
    WriteLine "<EN:LabAccreditation>"
    WriteLine "<EN:LabAccreditationIdentifier>000</EN:LabAccreditationIdentifier>"
    WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
    WriteLine "</EN:LabAccreditation>"
    WriteLine "</EN:LabIdentification>"
' === END Output Lab ID

' === Loop through samples
'not done
SamplesLoop:
    If Not exportSamples Or TableIsEmpty(SamplesTable) Then GoTo ResultsLoop
    
    Dim tbl As ListObject
    Set tbl = Range(SamplesTable).ListObject
    
    WriteLine "<EN:Sample>"
    
    ' TODO: LOOKUPS
    
    Dim row As Range
    For Each row In tbl.DataBodyRange.Rows
        WriteLine "<EN:SampleIdentification>"
        
        WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
        WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
        WriteLine CreateElement("EN:AdditionalSampleIndicator", CellValue(tbl, row, "Replacement"))
        WriteLine CreateElement("EN:PWSFacilityIdentifier", CellValue(tbl, row, "WSF State Assigned ID"))
        WriteLine CreateElement("EN:SampleRuleCode", "TC")
        WriteLine CreateElement("EN:ComplianceSampleIndicator", CellValue(tbl, row, "For Compliance"))
        WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
        WriteLine CreateElement("EN:SampleCollectionEndTime", CellTimeValue(tbl, row, "Sample Collection Time"))
        WriteLine CreateElement("EN:SampleMonitoringTypeCode", CellValue(tbl, row, "Sample Type"))
        WriteLine CreateElement("EN:SampleLaboratoryReceiptDate", CellDateValue(tbl, row, "Lab Receipt Date"))
        WriteLine WrapElement("SampleCollector", CreateElement("EN:IndividualFullName", CellValue(tbl, row, "Sample Collector Full Name")))
        
        WriteLine SpecializedMeasurement(CellValue(tbl, row, "Free Chlorine Residual (mg/L)"), "FreeChlorineResidual")
        WriteLine SpecializedMeasurement(CellValue(tbl, row, "Total Chlorine Residual (mg/L)"), "TotalChlorineResidual")
        
        If CellValue(tbl, row, "Sample Type") = "Repeat" Then
            WriteLine "<EN:OriginalSampleIdentification>"
            WriteLine CreateElement("EN:OriginalSampleIdentifier", CellValue(tbl, row, "Original Lab Sample ID"))
            WriteLine CreateElement("EN:OriginalSampleCollectionDate", CellDateValue(tbl, row, "Original Sample Collection Date"))
            WriteLine "<EN:OriginalSampleLabAccreditation>"
            WriteLine "<EN:LabAccreditationIdentifier>000</EN:LabAccreditationIdentifier>"
            WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
            WriteLine "</EN:OriginalSampleLabAccreditation>"
            WriteLine "</EN:OriginalSampleIdentification>"
        End If
        
        WriteLine "</EN:SampleIdentification>"
        
        WriteLine "<EN:SampleLocationIdentification>"
        WriteLine CreateElement("EN:SampleLocationIdentifier", CellValue(tbl, row, "Sampling Point ID"))
        WriteLine CreateElement("EN:SampleRepeatLocationCode", CellValue(tbl, row, "Repeat Location"))
        WriteLine "</EN:SampleLocationIdentification>"
    Next

    WriteLine "</EN:Sample>"
' === Loop through samples END
    
    
' === Loop through results
'not done
ResultsLoop:
    If Not exportResults Or TableIsEmpty(ResultsTable) Then GoTo Coda
    
    WriteLine "<EN:SampleAnalysisResults>"

    
    WriteLine "</EN:SampleAnalysisResults>"
' === END Loop through results
    
    
' === Close document
Coda:
    WriteLine "</EN:LabReport>"
    WriteLine "</EN:Submission>"
    WriteLine "</EN:eDWR>"
' === END Close document
    
My_Exit:
    ' Close file
    If Not closing Then
        closing = True
        Close #1
    End If
    
    Exit Sub

ErrHandler:
    If production Then MsgBox Err.Description
    Debug.Print Err.Description
    Resume My_Exit

End Sub

''' File functions
Private Sub WriteLine(line As String)
    If production Then
        Print #1, line
    Else
        Debug.Print line
    End If
End Sub

''' Complex data types

Private Function SpecializedMeasurement(value As Variant, typeCode As String) As String
    If value = Empty Then Exit Function
    
    Dim children As New Collection
    children.Add CreateElement("EN:MeasurementValue", CStr(value))
    children.Add CreateElement("EN:MeasurementSignificantDigit", GetSigFigs(value))
    children.Add CreateElement("EN:SpecializedMeasurementTypeCode", typeCode)
    
    SpecializedMeasurement = CreateParentElement("EN:SpecializedMeasurement", children)
End Function


''' Value functions
Function GetSigFigs(value As Variant) As Integer
    Dim val As String
    val = CStr(CDec(value))
    
    GetSigFigs = Len(val) - InStr(val, ".")
End Function

''' Macro
Sub ExportAllData()
    production = True
    exportSamples = True
    exportResults = True
    
    GenerateXmlDocument
End Sub

Sub TEST()
    WriteLine "===TEST==="
    
    Dim a As Variant
    
    a = CDec(3.3)
    WriteLine CStr(a)
    
    WriteLine GetSignificantDigits(a)
    WriteLine GetSignificantDigits(3.05)
    WriteLine GetSignificantDigits(13.05)
    WriteLine GetSignificantDigits(0.05)
    WriteLine GetSignificantDigits(0.05)
    
End Sub