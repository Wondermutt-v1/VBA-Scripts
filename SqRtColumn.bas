Attribute VB_Name = "RMSCalculation"
Sub SqColumn()

rowCnt = Cells(Rows.Count, 1).End(xlUp).Row
 Cells(3, 3).Value = rowCnt

 

For i = 3 To rowCnt
If ActiveSheet.Cells(i, 2).Value < 10000 Then


ActiveSheet.Cells(i, 3).Value = ActiveSheet.Cells(i, 2).Value * ActiveSheet.Cells(i, 2).Value
 
Else

Rows(i).Delete
i = i - 1

End If

Next


End Sub
