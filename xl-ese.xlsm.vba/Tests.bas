Option Explicit
Option Private Module

''' Globals

Sub Test_GetSigFigs()
    Debug.Print "=== GetSigFigs ==="
    Dim a As Variant
    a = CDec(3.3)
    
    Debug.Print "1", GetSigFigs(a)
    Debug.Print "1", GetSigFigs(CStr(a))
    
    Debug.Print "2", GetSigFigs(3.05)
    Debug.Print "2", GetSigFigs(13.05)
    Debug.Print "2", GetSigFigs(0.05)
End Sub

''' XmlFunctions

Sub Test_ReplaceEntities()
    Debug.Print "=== ReplaceEntities ==="
    Debug.Print "1&amp;2", ReplaceEntities("1&2")
    Debug.Print "&amp;", ReplaceEntities("&")
    Debug.Print "&lt;", ReplaceEntities("<")
    Debug.Print "&gt;", ReplaceEntities(">")
    Debug.Print "&quot;", ReplaceEntities("""")
    Debug.Print "&apos;", ReplaceEntities("'")
    Debug.Print "&amp;&gt;", ReplaceEntities("&>")
    Debug.Print "&gt;&amp;", ReplaceEntities(">&")
End Sub


Sub Test_CreateElement()
    Debug.Print "=== CreateElement ==="
    Debug.Print "<a>b</a>", CreateElement("a", "b")
    Debug.Print "[empty]", CreateElement("a", "")
    Debug.Print "<a></a>", CreateElement("a", "", True)
    Debug.Print "[empty]", CreateElement("a", Empty)
    Debug.Print "<a></a>", CreateElement("a", Empty, True)
End Sub

Sub Test_WrapElement()
    Debug.Print "=== WrapElement ==="
    Dim expected As String
    
    expected = "<a>" & vbNewLine & "<b />" & vbNewLine & "</a>"
    Debug.Print "expected"
    Debug.Print expected
    Debug.Print "actual"
    Debug.Print WrapElement("a", "<b />")
    
    expected = "[empty]"
    Debug.Print "expected"
    Debug.Print expected
    Debug.Print "actual"
    Debug.Print WrapElement("a", "")
    
    expected = "<a>" & vbNewLine & vbNewLine & "</a>"
    Debug.Print "expected"
    Debug.Print expected
    Debug.Print "actual"
    Debug.Print WrapElement("a", "", True)
End Sub

Sub Test_CreateParentElement()
    Debug.Print "=== CreateParentElement ==="
    Dim expected As String
    Dim children As New Collection
    
    expected = "<a>" & vbNewLine & "</a>"
    Debug.Print "expected"
    Debug.Print expected
    Debug.Print "actual"
    Debug.Print CreateParentElement("a", children)
    
    children.Add ("<b>1</b>")
    children.Add ("<b>2</b>")
    
    expected = "<a>" & vbNewLine & "<b>1</b>" & "<b>2</b>" & "</a>"
    Debug.Print "expected"
    Debug.Print expected
    Debug.Print "actual"
    Debug.Print CreateParentElement("a", children)
End Sub

''' TableFunctions

Sub Test_TableIsEmpty()
    Debug.Print "=== Test_TableIsEmpty ==="
    Debug.Print "false", TableIsEmpty("YesNoTable")
    Debug.Print "true", TableIsEmpty("EmptyTable")
End Sub

Sub Test_CellValue()
    Debug.Print "=== Test_CellValue ==="
    Dim tbl As ListObject
    Set tbl = Range("YesNoTable").ListObject

    Debug.Print "Y", CellValue(tbl, tbl.DataBodyRange.Rows, "Code")
End Sub

Sub Test_Lookup()
    Debug.Print "=== Test_Lookup ==="
    Debug.Print "Y", Lookup("Yes", "YesNoTable")
    Debug.Print "N", Lookup("No", "YesNoTable")
    Debug.Print "[empty]", Lookup("", "YesNoTable")
    Debug.Print "3014", Lookup("E. Coli", "AnalyteTable")
    Debug.Print "RT", Lookup("Routine", "SampleTypesTable")

    Debug.Print "An error should print below:"
    Debug.Print "[ERROR]", Lookup("Maybe", "YesNoTable")
End Sub

''' Files

Sub Test_Paths()
    Debug.Print Application.ThisWorkbook.Path
    Debug.Print Application.ThisWorkbook.name
    Debug.Print Application.ThisWorkbook.FullName
    Debug.Print Application.ThisWorkbook.FileFormat
    Debug.Print "Exists: " & Dir(Application.ThisWorkbook.FullName)
    Debug.Print "Not exists: " & Dir(Application.ThisWorkbook.FullName & ".nope")
End Sub

Sub Test_FileSaveDialog()
    Dim initPath As String
    initPath = Replace(Application.ThisWorkbook.FullName, ".xlsm", ".xml")
    Debug.Print Application.GetSaveAsFilename(initPath, "XML Files (*.xml), *.xml")
End Sub