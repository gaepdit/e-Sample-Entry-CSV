Option Explicit

Dim ExportSamples As Boolean
Dim ExportResults As Boolean
Dim Ready As Boolean

''' Export functions

Private Sub DebugExportAllData()
    ' Set params
    production = False
    ExportSamples = True
    ExportResults = False
    Ready = True
    
    Debug.Print
    Debug.Print Now
    Debug.Print "==="
   
    If GenerateXmlDocument Then
        Debug.Print "=== Success"
    Else
        Debug.Print "=== Error"
    End If
End Sub

Sub ExportAllData()
    ' Set params
    production = True
    ExportSamples = True
    ExportResults = True
    Ready = True
    
    If GenerateXmlDocument Then
        AlertMessage "File successfully created"
    Else
        AlertError "Error saving file"
    End If
End Sub

''' Generate data
Private Function GenerateXmlDocument() As Boolean
    ' Check params are set
    If Not Ready Then
        Debug.Print "!!! Params not set !!!"
        Exit Function
    End If

    ' Ensure data exists
    If Not ExportSamples And Not ExportResults Then Exit Function
    If Not ExportSamples And TableIsEmpty("ResultsDataTable") Then Exit Function
    If Not ExportResults And TableIsEmpty("SamplesDataTable") Then Exit Function
    
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
    
' --- END Create file

' === Start document
    WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    WriteLine "<EN:eDWR xmlns:EN=""urn:us:net:exchangenetwork"" xmlns:SDWIS=""http://www.epa.gov/sdwis"" xmlns:ns2=""http://www.epa.gov/xml"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
    WriteLine "<EN:Submission EN:submissionFileCreatedDate=""" & Format(Now, "yyyy-mm-dd") & """>"
    WriteLine "<EN:LabReport>"
' --- END Start document

' === Output LabIdentification
    WriteLine "<EN:LabIdentification>"
    WriteLine "<EN:LabAccreditation>"
    WriteLine "<EN:LabAccreditationIdentifier>000</EN:LabAccreditationIdentifier>"
    WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
    WriteLine "</EN:LabAccreditation>"
    WriteLine "</EN:LabIdentification>"
' --- END Output Lab ID

' === Loop through samples
SamplesLoop:
    If Not ExportSamples Or TableIsEmpty("SamplesDataTable") Then GoTo ResultsLoop
    
    Dim tbl As ListObject
    Set tbl = Range("SamplesDataTable").ListObject
    
    Dim row As Range
    For Each row In tbl.DataBodyRange.Rows
        WriteLine "<EN:Sample>"
        
        WriteLine "<EN:SampleIdentification>"
        WriteLine CreateElement("EN:StateSampleIdentifier", CellValue(tbl, row, "State Sample Number"))
        WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
        WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
        WriteLine CreateElement("EN:PWSFacilityIdentifier", CellValue(tbl, row, "WSF State Assigned ID"))
        WriteLine CreateElement("EN:SampleRuleCode", "TC")
        WriteLine CreateElement("EN:SampleMonitoringTypeCode", Lookup(CellValue(tbl, row, "Sample Type"), "SampleTypesTable"))
        WriteLine CreateElement("EN:ComplianceSampleIndicator", Lookup(CellValue(tbl, row, "For Compliance"), "YesNoTable"))
        WriteLine CreateElement("EN:AdditionalSampleIndicator", Lookup(CellValue(tbl, row, "Replacement"), "YesNoTable"))
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
        WriteLine WrapElement("EN:SampleCollector", CreateElement("ns2:IndividualFullName", CellValue(tbl, row, "Sample Collector Full Name")))
        WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
        WriteLine CreateElement("EN:SampleCollectionEndTime", CellTimeValue(tbl, row, "Sample Collection Time"))
        WriteLine CreateElement("EN:SampleLaboratoryReceiptDate", CellDateValue(tbl, row, "Lab Receipt Date"))
        WriteLine SpecializedMeasurement(CellValue(tbl, row, "Free Chlorine Residual (mg/L)"), "FreeChlorineResidual")
        WriteLine SpecializedMeasurement(CellValue(tbl, row, "Total Chlorine Residual (mg/L)"), "TotalChlorineResidual")
        WriteLine "</EN:SampleIdentification>"
        
        WriteLine "<EN:SampleLocationIdentification>"
        WriteLine CreateElement("EN:SampleLocationIdentifier", CellValue(tbl, row, "Sampling Point ID"))
        If CellValue(tbl, row, "Sample Type") = "Repeat" Then
            WriteLine CreateElement("EN:SampleRepeatLocationCode", Lookup(CellValue(tbl, row, "Repeat Location"), "RepeatLocationsTable"))
        End If
        WriteLine "</EN:SampleLocationIdentification>"
        WriteLine "</EN:Sample>"
    Next
' --- END Loop through samples

' === Loop through results
'not done
ResultsLoop:
    If Not ExportResults Or TableIsEmpty("ResultsDataTable") Then GoTo Coda
    
    WriteLine "<EN:SampleAnalysisResults>"

    
    WriteLine "</EN:SampleAnalysisResults>"
' --- END Loop through results
    
    
' === Close document
Coda:
    WriteLine "</EN:LabReport>"
    WriteLine "</EN:Submission>"
    WriteLine "</EN:eDWR>"
    GenerateXmlDocument = True
' === END Close document
    
My_Exit:
    ' Close file
    If Not closing Then
        closing = True
        Close #1
    End If
    
    Exit Function

ErrHandler:
    If production Then MsgBox Err.Description
    Debug.Print Err.Description
    Resume My_Exit

End Function

''' Complex data types

Private Function SpecializedMeasurement(value As Variant, typeCode As String) As String
    If value = Empty Then Exit Function
    
    Dim children As New Collection
    children.Add CreateElement("EN:MeasurementValue", CStr(value))
    children.Add CreateElement("EN:MeasurementSignificantDigit", GetSigFigs(value))
    children.Add CreateElement("EN:SpecializedMeasurementTypeCode", typeCode)
    
    SpecializedMeasurement = CreateParentElement("EN:SpecializedMeasurement", children)
End Function