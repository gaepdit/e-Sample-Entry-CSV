Sub SaveTableAsCsv()
    
    Dim xFile As Variant
    Dim xFileString As String

    Dim fileName As String
    fileName = "E:\projects\ese-csv\export.csv"
    
    ThisWorkbook.Sheets("Samples").ListObjects("SamplesTable").DataBodyRange.Copy
    
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
    'Range("E:F").NumberFormat = "yyyy-mm-dd"
    
    ActiveWorkbook.SaveAs fileName:=fileName, FileFormat:=xlCSVUTF8, CreateBackup:=False
    ActiveWorkbook.Close
    
End Sub
