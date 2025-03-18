VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub hyperlink_title()
    Dim i As Integer
    Dim wsh As Worksheet
    Set wsh = ThisWorkbook.ActiveSheet

    i = 2
    With wsh
    While ActiveSheet.Cells(i, 2) <> ""

        ActiveSheet.Hyperlinks.Add Anchor:=.Cells(i, 1), Address:=.Cells(i, 2).Value, _
        TextToDisplay:=.Cells(i, 1).Value2

        i = i + 1

    Wend
    End With
    

End Sub

