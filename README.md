# VBA_WordTable
''''Add Table to Word Document
Sub VerySimpleTableAdd()
    Dim oTable As Table
    Set oTable = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=3, NumColumns:=3)
End Sub

''''Select Table in Word
Sub SelectTable()
'selects first table in active doc
    If ActiveDocument.Tables.Count > 0 Then    'to avoid errors we check if any table exists in active doc
        ActiveDocument.Tables(1).Select
    End If
End Sub

'''Loop Through all Cells in a Table
Sub TableCycling()
' loop through all cells in table
    Dim nCounter As Long    ' this will be writen in all table cells
    Dim oTable As Table
    Dim oRow As Row
    Dim oCell As Cell

    ActiveDocument.Range.InsertParagraphAfter    'just makes new para athe end of doc, Table will be created here
    Set oTable = ActiveDocument.Tables.Add(Range:=ActiveDocument.Paragraphs.Last.Range, NumRows:=3, NumColumns:=3)    'create table and asign it to variable
    For Each oRow In oTable.Rows    ' outher loop goes through rows
        For Each oCell In oRow.Cells    'inner loop goes
            nCounter = nCounter + 1    'increases the counter
            oCell.Range.Text = nCounter    'writes counter to the cell
        Next oCell
    Next oRow

    'display result from cell from second column in second row
    Dim strTemp As String
    strTemp = oTable.Cell(2, 2).Range.Text
    MsgBox strTemp
End Sub

'Create Word Table From Excel File
Sub MakeTablefromExcelFile()
'advanced
    Dim oExcelApp, oExcelWorkbook, oExcelWorksheet, oExcelRange
    Dim nNumOfRows As Long
    Dim nNumOfCols As Long
    Dim strFile As String

    Dim oTable As Table    'word table
    Dim oRow As Row    'word row
    Dim oCell As Cell    'word table cell
    Dim x As Long, y As Long    'counter for loops

    strFile = "c:\Users\Nenad\Desktop\BookSample.xlsx"    'change to actual path
    Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible = True
    Set oExcelWorkbook = oExcelApp.Workbooks.Open(strFile)    'open workbook and asign it to variable
    Set oExcelWorksheet = oExcelWorkbook.Worksheets(1)    'asign first worksheet to variable
    Set oExcelRange = oExcelWorksheet.Range("A1:C8")
    nNumOfRows = oExcelRange.Rows.Count
    nNumOfCols = oExcelRange.Columns.Count

    ActiveDocument.Range.InsertParagraphAfter    'just makes new para athe end of doc, Table will be created here
    Set oTable = ActiveDocument.Tables.Add(Range:=ActiveDocument.Paragraphs.Last.Range, NumRows:=nNumOfRows, NumColumns:=nNumOfCols)    'create table and asign it to variable
    '***real deal, table gets filled here
    For x = 1 To nNumOfRows
        For y = 1 To nNumOfCols
            oTable.Cell(x, y).Range.Text = oExcelRange.Cells(x, y).Value
        Next y
    Next x
    '***
    oExcelWorkbook.Close False
    oExcelApp.Quit
    With oTable.Rows(1).Range    'we can now apply some beautiness to our table :)
        .Shading.Texture = wdTextureNone
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColor = wdColorYellow
    End With
End Sub
