Attribute VB_Name = "HyperlinkTitle"
Sub hyperlink_title()
    Dim i As Integer
    Dim wksh As Worksheet
    Set wksh = ThisWorkbook.ActiveSheet

    i = 2
    With wksh
    While ActiveSheet.Cells(i, 2) <> ""

        ActiveSheet.Hyperlinks.Add Anchor:=.Cells(i, 1), Address:=.Cells(i, 2).Value, _
        TextToDisplay:=.Cells(i, 1).Value2

        i = i + 1

    Wend
    End With
    
rowcnt = Cells(Rows.Count, 1).End(xlUp).Row
'Range(Cells(7, 3), Cells(growCnt, 4)).ClearFormats
Range(Cells(2, 4), Cells(rowcnt, 7)).NumberFormat = "d-mmm-yy"
'Range(Cells(7, 6), Cells(growcnt, 6)).NumberFormat = "General"

End Sub



Sub StandardFormat()

rowcnt = Cells(Rows.Count, 1).End(xlUp).Row
'Range(Cells(7, 3), Cells(growCnt, 4)).ClearFormats
Range(Cells(2, 4), Cells(rowcnt, 7)).NumberFormat = "d-mmm-yy"
'Range(Cells(7, 6), Cells(growcnt, 6)).NumberFormat = "General"

End Sub
