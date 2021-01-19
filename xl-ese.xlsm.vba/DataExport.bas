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
    
    WriteLine "<EN:LabReport>" ' :LabReportDataType
' --- END Start document

' === Output LabIdentification
    WriteLine "<EN:LabIdentification>" ' :LabIdentificationDataType
    WriteLine "<EN:LabAccreditation>"
    WriteLine "<EN:LabAccreditationIdentifier>000</EN:LabAccreditationIdentifier>"
    WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
    WriteLine "</EN:LabAccreditation>"
    WriteLine "</EN:LabIdentification>"
' --- END Output Lab ID

' === Declare table variables
    Dim tbl As ListObject
    Dim row As Range
' --- END Declare variables

' === Loop through samples
SamplesLoop:
    If Not ExportSamples Or TableIsEmpty("SamplesDataTable") Then GoTo ResultsLoop
    
    Set tbl = Range("SamplesDataTable").ListObject
    For Each row In tbl.DataBodyRange.Rows
        WriteLine "<EN:Sample>" ' :SampleDataType
        
        WriteLine "<EN:SampleIdentification>" ' :SampleIdentificationDataType
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
        WriteLine SpecializedMeasurement("EN:SpecializedMeasurement", CellValue(tbl, row, "Free Chlorine Residual (mg/L)"), "FreeChlorineResidual")
        WriteLine SpecializedMeasurement("EN:SpecializedMeasurement", CellValue(tbl, row, "Total Chlorine Residual (mg/L)"), "TotalChlorineResidual")
        WriteLine "</EN:SampleIdentification>"
        
        WriteLine "<EN:SampleLocationIdentification>" ' :SampleLocationIdentificationDataType
        WriteLine CreateElement("EN:SampleLocationIdentifier", CellValue(tbl, row, "Sampling Point ID"))
        If CellValue(tbl, row, "Sample Type") = "Repeat" Then
            WriteLine CreateElement("EN:SampleRepeatLocationCode", Lookup(CellValue(tbl, row, "Repeat Location"), "RepeatLocationsTable"))
        End If
        WriteLine "</EN:SampleLocationIdentification>"
        
        WriteLine "</EN:Sample>"
    Next
    
    row = Nothing
    tbl = Nothing
' --- END Loop through samples

' === Loop through results
'not done
ResultsLoop:
    If Not ExportResults Or TableIsEmpty("ResultsDataTable") Then GoTo Coda
    
    Set tbl = Range("ResultsDataTable").ListObject
    For Each row In tbl.DataBodyRange.Rows
        WriteLine "<EN:SampleAnalysisResults>" ' :SampleAnalysisResults

        WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
        WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
        WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
        
        WriteLine "<EN:LabAnalysisIdentification>" ' :LabAnalysisIdentificationDataType
        WriteLine "<EN:LabAccreditation>" ' :LabAccreditationDataType
        WriteLine "<EN:LabAccreditationIdentifier>000</EN:LabAccreditationIdentifier>"
        WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
        WriteLine "</EN:LabAccreditation>"
        WriteLine "<EN:SampleAnalyticalMethod>" ' :MethodDataType
        WriteLine CreateElement("EN:MethodIdentifier", CellValue(tbl, row, "Analytical Method"))
        WriteLine "</EN:SampleAnalyticalMethod>"
        WriteLine SpecializedMeasurement("EN:SampleAnalyzedMeasure", CellValue(tbl, row, "Volume Analyzed"))
        WriteLine CreateElement("EN:AnalysisStartDate", CellDateValue(tbl, row, "Analysis Start Date"))
        WriteLine CreateElement("EN:AnalysisStartTime", CellTimeValue(tbl, row, "Analysis Start Time"))
        WriteLine CreateElement("EN:AnalysisEndDate", CellDateValue(tbl, row, "Analysis End Date"))
        WriteLine CreateElement("EN:AnalysisEndTime", CellTimeValue(tbl, row, "Analysis End Time"))
        WriteLine "</EN:LabAnalysisIdentification>"
        
        WriteLine "<EN:AnalyteIdentification>" ' :AnalyteIdentificationDataType
        WriteLine "</EN:AnalyteIdentification>"

        WriteLine "<EN:AnalysisResult>" ' :AnalysisResultDataType
        WriteLine "</EN:AnalysisResult>"

        WriteLine "<EN:QAQCSummary>" ' :QAQCSummaryDataType
        WriteLine "</EN:QAQCSummary>"

        WriteLine "</EN:SampleAnalysisResults>"
    Next
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

Private Function SpecializedMeasurement(tag As String, value As Variant, Optional typeCode As String) As String
    If value = Empty Then Exit Function
    
    Dim children As New Collection
    children.Add CreateElement("EN:MeasurementValue", CStr(value))
    children.Add CreateElement("EN:MeasurementSignificantDigit", GetSigFigs(value))
    If typeCode <> Empty Then
        children.Add CreateElement("EN:SpecializedMeasurementTypeCode", typeCode)
    End If
    
    SpecializedMeasurement = CreateParentElement(tag, children)
End Function