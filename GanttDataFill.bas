Attribute VB_Name = "GanttDataFill"
Function sysBreakdown(RqdSheet, SrcSheet, subSysName)

Dim sourceSheet As Worksheet
Dim balerSheet As Worksheet
Dim powerTrainSheet As Worksheet
Dim engineSheet As Worksheet
Dim cottonSheet As Worksheet
Dim TMSheet As Worksheet
Dim Gantt As Worksheet
'Ranges
'Dim SysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range




'initialize variables
Set Gantt = Sheets(RqdSheet)
Set sourceSheet = Sheets(SrcSheet)



rowCnt = sourceSheet.Cells(Rows.Count, 2).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

'Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, bcolCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))
growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row


If Len(sourceSheet.Cells(5, 2)) <> 0 Then
growcnt = growcnt + 1
'import baler data
'growCnt = growCnt + 1

Gantt.Cells(growcnt, 2).Value = subSysName
Gantt.Cells(growcnt, 2).Font.Bold = True



growcnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'count the rows
Gantt.Activate

Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)
growcnt = growcnt + 1



'growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

'Crit.Copy 'Gantt.Range(9, 1)
'Gantt.Activate
'Cells(growCnt, 8).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues
Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues
Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues
growcnt = growcnt + 1
End If
End Function

Function CopyData()

Dim sourceSheet As Worksheet
Dim balerSheet As Worksheet
Dim powerTrainSheet As Worksheet
Dim engineSheet As Worksheet
Dim cottonSheet As Worksheet
Dim TMSheet As Worksheet
Dim Gantt As Worksheet
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range




'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Baler Tests")



rowCnt = sourceSheet.Cells(Rows.Count, 2).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

'Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, bcolCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

'import baler data
Gantt.Cells(7, 2).Value = "Baler"
Gantt.Cells(7, 2).Font.Bold = True



growcnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'count the rows
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)
growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues



End Function



Function CopyBaler()

Dim sourceSheet As Worksheet
Dim balerSheet As Worksheet
Dim powerTrainSheet As Worksheet
Dim engineSheet As Worksheet
Dim cottonSheet As Worksheet
Dim TMSheet As Worksheet
Dim Gantt As Worksheet
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range




'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Baler Tests")



rowCnt = sourceSheet.Cells(Rows.Count, 2).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

'Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, bcolCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

'import baler data
Gantt.Cells(7, 2).Value = "Baler"
Gantt.Cells(7, 2).Font.Bold = True



growcnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'count the rows
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)
growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues



End Function

Function CopyCtnSpfc()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Cotton Picker Specific")


growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Cotton"
Gantt.Cells(growcnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)


growcnt = growcnt + 1

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If


End Function

Function CopyCab()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Cab Tests")

growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Cab"
Gantt.Cells(growcnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growcnt = growcnt + 1
'sourceSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopyEngine()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Engine Tests")
'Set engineSheet = Sheets("Engine Tests")

growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Engine"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If



End Function

Function CopyChasis()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Chasis Tests")
'Set chasisSheet = Sheets("Chasis Tests")

growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate


If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Chassis"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopyPwrTrn()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Power Train Tests")
'Set powerTrainSheet = Sheets("Power Train Tests")

growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Power Train"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues
End If

End Function

Function CopyElec()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Electrical Tests")


growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Electrical Systems"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)



'MsgBox (growCnt)
growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopyHyd()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Hydraulic Tests")


growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Hydraulic Systems"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)


growcnt = growcnt + 1

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues
End If

End Function

Function CopySteer()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Steering Systems")


growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate

If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))



growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Steering Systems"
Gantt.Cells(growcnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

growcnt = growcnt + 1

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues

End If

End Function
Function CopyTtlVhcl()
'Ranges
Dim sysName As Range
Dim ID As Range
Dim Desc As Range
Dim BalRng As Range
Dim engineers As Range
Dim SchdStart As Range
Dim SchdFinish As Range
Dim Priority As Range
Dim Crit As Range
Dim Risk As Range
Dim Status As Range
Dim Tester As Range
Dim SPS As Range


'initialize variables
Set Gantt = Sheets("2024 planning")
Set sourceSheet = Sheets("Total Vehicle")


growcnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
rowCnt = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row  'count the rows
colCnt = sourceSheet.Cells(4, Columns.Count).End(xlToLeft).Column
sourceSheet.Activate
If Len(Cells(5, 1)) <> 0 Then
Set BalRng = sourceSheet.Range(Cells(4, 1), Cells(rowCnt, colCnt))
Set ID = sourceSheet.Range(Cells(5, 2), Cells(rowCnt, 2))
Set Desc = sourceSheet.Range(Cells(5, 11), Cells(rowCnt, 11))
Set engineers = sourceSheet.Range(Cells(5, 36), Cells(rowCnt, 36))
Set SchdStart = sourceSheet.Range(Cells(5, 20), Cells(rowCnt, 20))
Set SchdFinish = sourceSheet.Range(Cells(5, 21), Cells(rowCnt, 21))
Set Priority = sourceSheet.Range(Cells(5, 32), Cells(rowCnt, 32))
Set Crit = sourceSheet.Range(Cells(5, 24), Cells(rowCnt, 24))
Set Risk = sourceSheet.Range(Cells(5, 23), Cells(rowCnt, 23))
Set Status = sourceSheet.Range(Cells(5, 7), Cells(rowCnt, 7))
Set Tester = sourceSheet.Range(Cells(5, 28), Cells(rowCnt, 28))
Set SPS = sourceSheet.Range(Cells(5, 45), Cells(rowCnt, 45))

growcnt = growcnt + 1
'import baler data
Gantt.Cells(growcnt, 2).Value = "Total Vehicle"
Gantt.Cells(growcnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growcnt, 1), Cells(growcnt, 15)).Font.Color = RGB(255, 255, 255)

growcnt = growcnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 2).PasteSpecial xlPasteValues

SchdStart.Copy '3
Gantt.Activate
Cells(growcnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy '4
Gantt.Activate
Cells(growcnt, 4).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 5).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growcnt, 6).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growcnt, 13).PasteSpecial xlPasteValues

Crit.Copy '12
Gantt.Activate
Cells(growcnt, 12).PasteSpecial xlPasteValues

Status.Copy '14
Gantt.Activate
Cells(growcnt, 14).PasteSpecial xlPasteValues

Tester.Copy '15
Gantt.Activate
Cells(growcnt, 15).PasteSpecial xlPasteValues
End If

End Function
