Option Explicit

''' Data debug
Private Sub DebugData()
    ''' If Debugging is True, all data is written to the VBA Immediate window instead of a file
    Debugging = True
    
    Debug.Print
    Debug.Print Now
    Debug.Print "==="
    
    If ExportAllData Then
        Debug.Print "=== Success"
    Else
        Debug.Print "=== Error"
    End If
End Sub

''' Generate data
Function ExportAllData() As Boolean
    On Error GoTo ErrHandler:
    
' === Basic data checks
    If ThisWorkbook.Names("LabCertNumber").RefersToRange(1, 1) = Empty Then
        Range("LabCertNumber").Select
        AlertMessage "Enter the Lab Certification Number before exporting."
        Exit Function
    End If
    
    If TableIsEmpty("SamplesDataTable") Then
        Range("SamplesDataTable").Select
        AlertMessage "There is no data to export."
        Exit Function
    End If
    
' === Create file
    If Debugging Then GoTo StartDocument
    
    If Application.ThisWorkbook.Path = vbNullString Then
        ' The workbook hasn't been saved yet
        ' (only really possible if starting from a template)
        AlertMessage "Please save the spreadsheet first."
        Exit Function
    End If
    
    ' generate default file name/path
    Dim fPath As String, saveAsResult As Variant
    fPath = Replace(Application.ThisWorkbook.FullName, " ", "_")
    fPath = Replace(fPath, ".xlsm", "") & ".xml"
    
GetFilePath:
    ' request file name/path from user
    saveAsResult = Application.GetSaveAsFilename(fPath, "XML Files (*.xml), *.xml")
    
    If saveAsResult = False Or saveAsResult = "" Then
        Exit Function
    End If
    
    ' Spaces are not allowed in the filename
    fPath = Replace(saveAsResult, " ", "_")
        
    ' If file exists, verify whether to overwrite or try again
    If Dir(fPath) <> "" Then
        If vbNo = MsgBox(Dir(fPath) & " already exists." & vbNewLine & "Do you want to replace it?", vbYesNo + vbExclamation + vbDefaultButton2, "Confirm Save As") Then
            GoTo GetFilePath
        End If
    End If
    
    ' Open file for saving
    FileNum = FreeFile
    Open fPath For Output As #FileNum

' === Start XML document
StartDocument:
    WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    WriteLine "<!-- Generated by xl-ESE version " & Format(ThisWorkbook.Names("AppVersion").RefersToRange(1, 1), "yyyy-mm-dd") & "; Excel " & Application.Version & "; " & Application.OperatingSystem & " -->"
    WriteLine "<EN:eDWR xmlns:EN=""urn:us:net:exchangenetwork"" xmlns:SDWIS=""http://www.epa.gov/sdwis"" xmlns:ns2=""http://www.epa.gov/xml"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
    WriteLine "<EN:Submission EN:submissionFileCreatedDate=""" & Format(Now, "yyyy-mm-dd") & """>"
    WriteLine "<EN:LabReport>" ' :LabReportDataType

' === Output LabIdentification
    Dim LabCertNumber As String
    LabCertNumber = ThisWorkbook.Names("LabCertNumber").RefersToRange(1, 1)
        
    WriteLine "<EN:LabIdentification>" ' :LabIdentificationDataType
    WriteLine "<EN:LabAccreditation>"
    WriteLine "<EN:LabAccreditationIdentifier>" & LabCertNumber & "</EN:LabAccreditationIdentifier>"
    WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
    WriteLine "</EN:LabAccreditation>"
    WriteLine "</EN:LabIdentification>"

' === Declare variables for loops
    Dim tbl As ListObject
    Dim row As Range
    Set tbl = Range("SamplesDataTable").ListObject
    
' === Loop through samples
SamplesLoop:
    Dim sampleType As String
    For Each row In tbl.DataBodyRange.Rows
        If CellValue(tbl, row, "Lab Sample ID") <> "" Then
            WriteLine "<EN:Sample>" ' :SampleDataType
            
            WriteLine "<EN:SampleIdentification>" ' :SampleIdentificationDataType
            WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
            WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
            WriteLine CreateElement("EN:PWSFacilityIdentifier", "950")
            WriteLine CreateElement("EN:SampleRuleCode", "TC")
            
            sampleType = Lookup(CellValue(tbl, row, "Sampling Point Type/Location"), "SampleTypesTable")
            WriteLine CreateElement("EN:SampleMonitoringTypeCode", sampleType)
            
            If sampleType = "SP" Then
                WriteLine CreateElement("EN:ComplianceSampleIndicator", "N")
            Else
                WriteLine CreateElement("EN:ComplianceSampleIndicator", "Y")
            End If
            
            If sampleType = "RP" Then
                WriteLine "<EN:OriginalSampleIdentification>" ' :OriginalSampleIdentificationDataType
                WriteLine CreateElement("EN:OriginalSampleIdentifier", CellValue(tbl, row, "Original Lab Sample ID"))
                WriteLine CreateElement("EN:OriginalSampleCollectionDate", CellDateValue(tbl, row, "Original Sample Collection Date"))
                WriteLine "<EN:OriginalSampleLabAccreditation>" ' :LabAccreditationDataType
                WriteLine "<EN:LabAccreditationIdentifier>" & LabCertNumber & "</EN:LabAccreditationIdentifier>"
                WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
                WriteLine "</EN:OriginalSampleLabAccreditation>"
                WriteLine "</EN:OriginalSampleIdentification>"
            End If
            
            WriteLine WrapElement("EN:SampleCollector", CreateElement("ns2:IndividualFullName", CellValue(tbl, row, "Sample Collector Full Name")))
            WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
            WriteLine CreateElement("EN:SampleCollectionEndTime", CellTimeValue(tbl, row, "Sample Collection Time"))
            WriteLine CreateElement("EN:SampleLaboratoryReceiptDate", CellDateValue(tbl, row, "Lab Receipt Date"))
            WriteLine SpecializedMeasurement("EN:SpecializedMeasurement", CellValue(tbl, row, "Free Chlorine Residual (mg/L)"), "FreeChlorineResidual")
            WriteLine "</EN:SampleIdentification>"
            
            WriteLine "<EN:SampleLocationIdentification>" ' :SampleLocationIdentificationDataType
            WriteLine CreateElement("EN:SampleLocationIdentifier", Lookup(CellValue(tbl, row, "Sampling Point Type/Location"), "SampleTypesTable", 3))
            
            If sampleType = "RP" Then
                WriteLine CreateElement("EN:SampleRepeatLocationCode", Lookup(CellValue(tbl, row, "Sampling Point Type/Location"), "SampleTypesTable", 4))
            End If
            
            WriteLine CreateElement("EN:SampleLocationCollectionAddress", CellValue(tbl, row, "Collection Address"))
            WriteLine "</EN:SampleLocationIdentification>"
            WriteLine "</EN:Sample>"
        End If
    Next
    
    Set row = Nothing

' === Loop through results
ResultsLoop:
    For Each row In tbl.DataBodyRange.Rows
        If CellValue(tbl, row, "Lab Sample ID") <> "" Then
            WriteLine "<EN:SampleAnalysisResults>" ' :SampleAnalysisResults
    
            WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
            WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
            WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
            
            WriteLine "<EN:LabAnalysisIdentification>" ' :LabAnalysisIdentificationDataType
            WriteLine "<EN:LabAccreditation>" ' :LabAccreditationDataType
            WriteLine "<EN:LabAccreditationIdentifier>" & LabCertNumber & "</EN:LabAccreditationIdentifier>"
            WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
            WriteLine "</EN:LabAccreditation>"
            WriteLine "<EN:SampleAnalyticalMethod>" ' :MethodDataType
            WriteLine CreateElement("EN:MethodIdentifier", CellValue(tbl, row, "Analytical Method"))
            WriteLine "</EN:SampleAnalyticalMethod>"
            WriteLine UnitMeasurement("EN:SampleAnalyzedMeasure", 100, "ML")
            WriteLine CreateElement("EN:AnalysisStartDate", CellDateValue(tbl, row, "Analysis Start Date"))
            WriteLine CreateElement("EN:AnalysisStartTime", CellTimeValue(tbl, row, "Start Time"))
            WriteLine CreateElement("EN:AnalysisEndDate", CellDateValue(tbl, row, "End Date"))
            WriteLine CreateElement("EN:AnalysisEndTime", CellTimeValue(tbl, row, "End Time"))
            WriteLine "</EN:LabAnalysisIdentification>"
            
            WriteLine "<EN:AnalyteIdentification>" ' :AnalyteIdentificationDataType
            WriteLine CreateElement("EN:AnalyteCode", "3100") ' : AnalyteCodeDataType
            WriteLine "</EN:AnalyteIdentification>"
    
            WriteLine "<EN:AnalysisResult>" ' :AnalysisResultDataType
            WriteLine "<EN:Result>" ' :MeasurementDataType
            
            WriteLine CreateElement("EN:MeasurementQualifier", Lookup(CellValue(tbl, row, "Presence"), "PresenceTable"))
            WriteLine "</EN:Result>"
            WriteLine CreateElement("EN:ResultStateNotificationDate", FormatDate(Date))
            WriteLine "</EN:AnalysisResult>"
    
            WriteLine "<EN:QAQCSummary>" ' :QAQCSummaryDataType
            WriteLine "<EN:DataQualityCode>A</EN:DataQualityCode>"
            WriteLine "</EN:QAQCSummary>"
    
            WriteLine "</EN:SampleAnalysisResults>"
            
            If CellValue(tbl, row, "Presence") = "Present" Then
                WriteLine "<EN:SampleAnalysisResults>" ' :SampleAnalysisResults
        
                WriteLine CreateElement("EN:LabSampleIdentifier", CellValue(tbl, row, "Lab Sample ID"))
                WriteLine CreateElement("EN:PWSIdentifier", CellValue(tbl, row, "PWS Number"))
                WriteLine CreateElement("EN:SampleCollectionEndDate", CellDateValue(tbl, row, "Sample Collection Date"))
                
                WriteLine "<EN:LabAnalysisIdentification>" ' :LabAnalysisIdentificationDataType
                WriteLine "<EN:LabAccreditation>" ' :LabAccreditationDataType
                WriteLine "<EN:LabAccreditationIdentifier>" & LabCertNumber & "</EN:LabAccreditationIdentifier>"
                WriteLine "<EN:LabAccreditationAuthorityName>STATE</EN:LabAccreditationAuthorityName>"
                WriteLine "</EN:LabAccreditation>"
                WriteLine "<EN:SampleAnalyticalMethod>" ' :MethodDataType
                WriteLine CreateElement("EN:MethodIdentifier", CellValue(tbl, row, "Analytical Method E"))
                WriteLine "</EN:SampleAnalyticalMethod>"
                WriteLine UnitMeasurement("EN:SampleAnalyzedMeasure", 100, "ML")
                WriteLine CreateElement("EN:AnalysisStartDate", CellDateValue(tbl, row, "Start Date E"))
                WriteLine CreateElement("EN:AnalysisStartTime", CellTimeValue(tbl, row, "Start Time E"))
                WriteLine CreateElement("EN:AnalysisEndDate", CellDateValue(tbl, row, "End Date E"))
                WriteLine CreateElement("EN:AnalysisEndTime", CellTimeValue(tbl, row, "End Time E"))
                WriteLine "</EN:LabAnalysisIdentification>"
                
                WriteLine "<EN:AnalyteIdentification>" ' :AnalyteIdentificationDataType
                WriteLine CreateElement("EN:AnalyteCode", "3014") ' : AnalyteCodeDataType
                WriteLine "</EN:AnalyteIdentification>"
        
                WriteLine "<EN:AnalysisResult>" ' :AnalysisResultDataType
                WriteLine "<EN:Result>" ' :MeasurementDataType
                
                WriteLine CreateElement("EN:MeasurementQualifier", Lookup(CellValue(tbl, row, "Presence E"), "PresenceTable"))
                WriteLine "</EN:Result>"
                WriteLine CreateElement("EN:ResultStateNotificationDate", FormatDate(Date))
                WriteLine "</EN:AnalysisResult>"
        
                WriteLine "<EN:QAQCSummary>" ' :QAQCSummaryDataType
                WriteLine "<EN:DataQualityCode>A</EN:DataQualityCode>"
                WriteLine "</EN:QAQCSummary>"
        
                WriteLine "</EN:SampleAnalysisResults>"
            End If
        End If
    Next

    Set row = Nothing
    
' === Close document
Coda:
    WriteLine "</EN:LabReport>"
    WriteLine "</EN:Submission>"
    WriteLine "</EN:eDWR>"
    
    ExportAllData = True
    AlertMessage "File successfully created."
    
' === Close file and exit
My_Exit:
    Close #FileNum
    Exit Function

ErrHandler:
    AlertError Err.Description
    Resume My_Exit

End Function