Attribute VB_Name = "GanttFunction2024"

Function TestAlert2024()
Dim Gantt As Worksheet
Set Gantt = Sheets("2024 planning")
strt = "To Be Started"



' highlight late tests
rowCnt = Cells(Rows.Count, 14).End(xlUp).Row
theDate = Now()
For i = 8 To rowCnt
    If Len(Cells(i, 14)) > 0 Then
        
        'Note the tests that are still open 90 days after It should have ended
        If Cells(i, 15).Value <= theDate - 90 Then
        Cells(i, 14).Select
            With Selection
                .Style = "PIINNNKKK"
                .NumberFormat = "d-mmm-yy"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
            End With
                        Cells(i, 15).Select
            With Selection
                .Style = "PIINNNKKK"
                .NumberFormat = "d-mmm-yy"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
            End With
            End If
        'Note tests that are over 30 days late to start
        If Cells(i, 14).Value <= theDate - 30 And Cells(i, 12).Value = strt Then
            Cells(i, 14).Select
            With Selection
                .Style = "Bad"
                .NumberFormat = "d-mmm-yy"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
            End With
        Cells(i, 15).Select
            With Selection
                .Style = "Bad"
                .NumberFormat = "d-mmm-yy"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
            End With
        End If
    
        If Len(Cells(i, 14)) <> 0 And Len(Cells(i, 12)) = 0 Then
            Cells(i, 14).Select
            Selection.Font.Color = RGB(255, 205, 196)
            Selection.Interior.Color = RGB(255, 0, 0)
        End If
End If


Next

End Function

Function CopyBaler2024()

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



growCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'count the rows
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)
growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End Function

Function CopyCtnSpfc2024()
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
Gantt.Cells(growCnt, 2).Value = "Cotton"
Gantt.Cells(growCnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)


growCnt = growCnt + 1

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If


End Function

Function CopyCab2024()
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
Gantt.Cells(growCnt, 2).Value = "Cab"
Gantt.Cells(growCnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growCnt = growCnt + 1
'sourceSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopyEngine2024()
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
Gantt.Cells(growCnt, 2).Value = "Engine"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues


End If



End Function

Function CopyChasis2024()
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
Gantt.Cells(growCnt, 2).Value = "Chassis"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

'MsgBox (growCnt)
growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues


End If

End Function

Function CopyPwrTrn2024()
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
Gantt.Cells(growCnt, 2).Value = "Power Train"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopyElec2024()
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
Gantt.Cells(growCnt, 2).Value = "Electrical Systems"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)



'MsgBox (growCnt)
growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues


End If

End Function

Function CopyHyd2024()
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
Gantt.Cells(growCnt, 2).Value = "Hydraulic Systems"
Gantt.Cells(growCnt, 2).Font.Bold = True
Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)


growCnt = growCnt + 1

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function

Function CopySteer2024()
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

ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function
Function CopyTtlVhcl2024()
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
Gantt.Cells(growCnt, 2).Value = "Total Vehicle"
Gantt.Cells(growCnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function


Function CopyLabs2024()
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
Gantt.Cells(growCnt, 2).Value = "Lab Tests"
Gantt.Cells(growCnt, 2).Font.Bold = True

Gantt.Activate
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Interior.Color = RGB(27, 95, 169)
Gantt.Range(Cells(growCnt, 1), Cells(growCnt, 15)).Font.Color = RGB(255, 255, 255)

growCnt = growCnt + 1
'chasisSheet.
ID.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 1).PasteSpecial xlPasteValues

Desc.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 2).PasteSpecial xlPasteValues

engineers.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 3).PasteSpecial xlPasteValues

Priority.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 5).PasteSpecial xlPasteValues

SPS.Copy
Gantt.Activate
Cells(growCnt, 7).PasteSpecial xlPasteValues

Crit.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 6).PasteSpecial xlPasteValues

Status.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 12).PasteSpecial xlPasteValues

Tester.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 13).PasteSpecial xlPasteValues

SchdStart.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 14).PasteSpecial xlPasteValues

SchdFinish.Copy 'Gantt.Range(9, 1)
Gantt.Activate
Cells(growCnt, 15).PasteSpecial xlPasteValues

End If

End Function

