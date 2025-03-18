Attribute VB_Name = "TimeOnOffCalculator"
Sub TimeOnOffCounter()

'First find the number of rows in the file.
'This will be used in the For statement
rowCnt = Cells(Rows.Count, 1).End(xlUp).Row
colCnt = Cells(1, Columns.Count).End(xlToLeft).Column
Cells(3, 3).Value = rowCnt
i = 0
tStart = 0
tStop = 0
timeOn = 3
'Start reading the values
For i = 3 To rowCnt

'Check for an on/off scenario.
'Detect a "rising edge"
If ActiveSheet.Cells(i, 2).Value > 0 And ActiveSheet.Cells(i - 1, 2).Value = 0 And ActiveSheet.Cells(i - 1, 2).Value <= 600 Then
tStart = ActiveSheet.Cells(i, 1).Value

'Check for Falling Edge
ElseIf ActiveSheet.Cells(i, 2).Value = 0 And ActiveSheet.Cells(i - 1, 2).Value > 0 And i <> 3 Then
tStop = ActiveSheet.Cells(i - 1, 1).Value

ActiveSheet.Cells(timeOn, colCnt + 4).Value = tStop - tStart
timeOn = timeOn + 1



Else


End If

Next
tmRows = Cells(Rows.Count, colCnt + 4).End(xlUp).Row
tTm = 0
For m = 3 To tmRows

tTm = tTm + ActiveSheet.Cells(m, colCnt + 4).Value
Next
ActiveSheet.Cells(3, colCnt + 5).Value = tTm
ActiveSheet.Cells(3, colCnt + 6).Value = tTm / 60
End Sub

