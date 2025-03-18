Attribute VB_Name = "GanttFunctions"
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
Set Gantt = Sheets("Schedule Planning")
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
'Gantt.Cells(7, 2).Value = "Baler"
'Gantt.Cells(7, 2).Font.Bold = True


Gantt.Activate
growCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'count the rows

'Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
'Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)
growCnt = growCnt + 2

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

Cells(growCnt + 1, 1).Activate






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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Cotton Picker Specific")


growCnt = Gantt.Cells(Rows.Count, 3).End(xlUp).Row
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

growCnt = growCnt + 1


Gantt.Activate

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Cab Tests")

growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1

Gantt.Activate

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Engine Tests")
'Set engineSheet = Sheets("Engine Tests")

growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1

Gantt.Activate

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues


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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Chasis Tests")
'Set chasisSheet = Sheets("Chasis Tests")

growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data
'Gantt.Cells(growCnt, 2).Value = "Chassis"
'Gantt.Cells(growCnt, 2).Font.Bold = True
'Gantt.Activate
'Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
'Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues


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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Power Train Tests")
'Set powerTrainSheet = Sheets("Power Train Tests")

growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data

Gantt.Activate


growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Electrical Tests")


growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data


'MsgBox (growCnt)
growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues


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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Hydraulic Tests")


growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data



growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Steering Systems")


growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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



growCnt = growCnt + 1
'import baler data
Gantt.Cells(growCnt, 2).Value = "Steering Systems"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Total Vehicle")


growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

End If

End Function



Function CopyLabs()
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
Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("Lab Tests")


growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
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

growCnt = growCnt + 1
'import baler data

growCnt = growCnt + 1

ID.Copy
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 8).PasteSpecial xlPasteValues

Status.Copy
Gantt.Activate
Cells(growCnt, 9).PasteSpecial xlPasteValues

SchdStart.Copy
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

SchdFinish.Copy
Gantt.Activate
Cells(growCnt, 4).PasteSpecial xlPasteValues

End If

End Function
