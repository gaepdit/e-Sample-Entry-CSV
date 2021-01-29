Option Explicit

Dim ExportSamples As Boolean
Dim ExportResults As Boolean
Dim Ready As Boolean

''' Data export functions

Private Sub DebugExportAllData()
    ''' Writes all data to the VBA Immediate window instead of a file
    ' Set params
    Production = False
    ExportSamples = True
    ExportResults = True
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
    Production = True
    ExportSamples = True
    ExportResults = True
    Ready = True
    
    GenerateXmlDocument
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
    If TableIsEmpty("SamplesDataTable") And TableIsEmpty("ResultsDataTable") Then
        AlertMessage "There is no data to export."
        Exit Function
    End If
    If Not ExportSamples And TableIsEmpty("ResultsDataTable") Then
        AlertMessage "There is no data to export."
        Exit Function
    End If
    If Not ExportResults And TableIsEmpty("SamplesDataTable") Then
        AlertMessage "There is no data to export."
        Exit Function
    End If
    
    ' For error handling
    Dim closing As Boolean
    closing = False
    
    On Error GoTo ErrHandler:
    
' === Initial data validation (?)
'not done
    
' === Create file
    If Application.ThisWorkbook.Path = vbNullString Then
        ' The workbook hasn't been saved yet
        ' (only really possible if starting from a template)
        AlertMessage "Please save the spreadsheet first."
        Exit Function
    End If
    
    ' generate default file name/path
    Dim fPath As String, saveAsResult As Variant
    fPath = Replace(Application.ThisWorkbook.FullName, ".xlsm", "") & ".xml"
    
GetFilePath:
    ' request file name/path from user
    saveAsResult = Application.GetSaveAsFilename(fPath, "XML Files (*.xml), *.xml")
    
    If saveAsResult = False Then
        Exit Function
    Else
        fPath = saveAsResult
    End If
    
    ' If file exists, verify whether to overwrite or try again
    If Dir(fPath) <> "" Then
        If vbNo = MsgBox(Dir(fPath) & " already exists." & vbNewLine & "Do you want to replace it?", vbYesNo + vbExclamation + vbDefaultButton2, "Confirm Save As") Then
            GoTo GetFilePath
        End If
    End If
    
    ' Open file for saving
    FileNum = FreeFile
    Open fPath For Output As #FileNum
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
        If CellValue(tbl, row, "Lab Sample ID") <> "" Then
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
                WriteLine "<EN:OriginalSampleIdentification>" ' :OriginalSampleIdentificationDataType
                WriteLine CreateElement("EN:OriginalSampleIdentifier", CellValue(tbl, row, "Original Lab Sample ID"))
                WriteLine CreateElement("EN:OriginalSampleCollectionDate", CellDateValue(tbl, row, "Original Sample Collection Date"))
                WriteLine "<EN:OriginalSampleLabAccreditation>" ' :LabAccreditationDataType
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
        End If
    Next
    
    Set row = Nothing
    Set tbl = Nothing
' --- END Loop through samples

' === Loop through results
ResultsLoop:
    If Not ExportResults Or TableIsEmpty("ResultsDataTable") Then GoTo Coda
    
    Set tbl = Range("ResultsDataTable").ListObject
    For Each row In tbl.DataBodyRange.Rows
        If CellValue(tbl, row, "Lab Sample ID") <> "" Then
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
            WriteLine UnitMeasurement("EN:SampleAnalyzedMeasure", Lookup(CellValue(tbl, row, "Volume Analyzed"), "VolumeTable"), "ML")
            WriteLine CreateElement("EN:AnalysisStartDate", CellDateValue(tbl, row, "Analysis Start Date"))
            WriteLine CreateElement("EN:AnalysisStartTime", CellTimeValue(tbl, row, "Analysis Start Time"))
            WriteLine CreateElement("EN:AnalysisEndDate", CellDateValue(tbl, row, "Analysis End Date"))
            WriteLine CreateElement("EN:AnalysisEndTime", CellTimeValue(tbl, row, "Analysis End Time"))
            WriteLine "</EN:LabAnalysisIdentification>"
            
            WriteLine "<EN:AnalyteIdentification>" ' :AnalyteIdentificationDataType
            WriteLine CreateElement("EN:AnalyteCode", Lookup(CellValue(tbl, row, "Analyte"), "AnalyteTable")) ' : AnalyteCodeDataType
            WriteLine "</EN:AnalyteIdentification>"
    
            WriteLine "<EN:AnalysisResult>" ' :AnalysisResultDataType
            WriteLine "<EN:Result>" ' :MeasurementDataType
            WriteLine CreateElement("EN:MeasurementQualifier", Lookup(CellValue(tbl, row, "Microbe Presence"), "PresenceTable"))
            If CellValue(tbl, row, "Microbe Presence") = "Present" And CellValue(tbl, row, "Result Count") <> Empty Then
                WriteLine CreateElement("EN:MeasurementValue", CellValue(tbl, row, "Result Count"))
                WriteLine CreateElement("EN:MeasurementUnit", CellValue(tbl, row, "per Volume")) ' Don't use lookup code for count volume units
                WriteLine CreateElement("EN:MicrobialResultCountTypeCode", Lookup(CellValue(tbl, row, "Units"), "CountUnitsTable"))
            End If
            WriteLine "</EN:Result>"
            WriteLine CreateElement("EN:ResultStateNotificationDate", CellDateValue(tbl, row, "State Notification Date"))
            WriteLine "</EN:AnalysisResult>"
    
            WriteLine "<EN:QAQCSummary>" ' :QAQCSummaryDataType
            WriteLine "<EN:DataQualityCode>A</EN:DataQualityCode>"
            WriteLine "</EN:QAQCSummary>"
    
            WriteLine "</EN:SampleAnalysisResults>"
        End If
    Next
' --- END Loop through results
    
' === Close document
Coda:
    WriteLine "</EN:LabReport>"
    WriteLine "</EN:Submission>"
    WriteLine "</EN:eDWR>"
    
    GenerateXmlDocument = True
    AlertMessage "File successfully created."
' --- END Close document
    
' === Close file
My_Exit:
    If Not closing Then
        closing = True
        Close #FileNum
    End If
' --- END Close file
    
    Exit Function

ErrHandler:
    AlertError Err.Description
    Resume My_Exit

End Function