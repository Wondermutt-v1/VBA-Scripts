Attribute VB_Name = "WOImporter"
Sub WOImport()

Application.ScreenUpdating = False

Dim importFile As String
Dim destLocation As Worksheet

Set destWB = ThisWorkbook
Set destLocation = ThisWorkbook.Worksheets("WO Data")

importFile = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx; *.xlsm; *.xls),*.xlsx;*.xlsm;*.xls", Title:="Choose the file containing CTS Data for import.", MultiSelect:=False)
If importFile = "False" Then Exit Sub

destLocation.Range(destLocation.Cells(5, 1), destLocation.Cells(10000, 100)).Clear 'Clears all data in range to be pasted into

Set wbSource = Workbooks.Open(importFile)
Set DataSource = wbSource.ActiveSheet

If IsError(Application.Match("ID", DataSource.Range("1:1"), 0)) Then
    MsgBox ("Data must contain CTS ID number with column header ID.")
    Exit Sub
End If

With DataSource
    lastSourceCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    IDCol = Application.Match("ID", .Range("1:1"), 0)
    lastRow = .Cells(Rows.Count, IDCol).End(xlUp).Row
    dataToCopy = .Range(Cells(1, 1), Cells(lastRow + 1, lastSourceCol)).Value
End With

With destLocation
    .Range(.Cells(4, 1), .Cells(4 + lastRow, lastSourceCol)).Value = dataToCopy
    .Range(.Cells(4, 1), .Cells(4 + lastRow, lastSourceCol)).WrapText = False
    .Range(.Cells(4, 1), .Cells(4 + lastRow, lastSourceCol)).HorizontalAlignment = xlLeft
End With

destLocation.Columns(IDCol).FormatConditions.Delete

'inputFormula = "=AND(ISNA(MATCH($B1,'Primary IVP'!B:B,0)),NOT(ISBLANK($B1)))"
'formatFormula = LocalizeFormula((inputFormula), destLocation)

With destLocation.Columns(IDCol).FormatConditions.Add(Type:=xlExpression, Formula1:=formatFormula)
    .Interior.Color = RGB(250, 200, 70)
End With

Application.DisplayAlerts = False
wbSource.Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True
   
Sheets("TR Data").Cells(1, 4).Value = "Last Updated:"
Sheets("TR Data").Cells(1, 5).Value = Now()
Sheets("TR Data").Cells(1, 5).NumberFormat = "dd-mmm-yy"

End Sub

