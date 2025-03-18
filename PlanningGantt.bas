Attribute VB_Name = "PlanningGantt"

Sub PlanGantt()


Dim Gantt As Worksheet
Dim sourceSheet As Worksheet
Dim balerSheet As Worksheet
Dim powerTrainSheet As Worksheet
Dim engineSheet As Worksheet
Dim cottonSheet As Worksheet
Dim chasisSheet As Worksheet
Dim TMSheet As Worksheet
Dim elctrcSheet As Worksheet
Dim hydSheet As Worksheet
Dim steerSheet As Worksheet
Dim TotlVhcl As Worksheet
Dim GanttTable As ListObject
Dim objTable As ListObject


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
Dim GanttCal As Range

Sheets("Power Train Tests").Visible = True
Sheets("Chasis Tests").Visible = True
Sheets("Baler Tests").Visible = True
Sheets("Engine Tests").Visible = True
Sheets("Cotton Picker Specific").Visible = True
Sheets("Cab Tests").Visible = True
Sheets("Electrical Tests").Visible = True
Sheets("Hydraulic Tests").Visible = True
Sheets("Steering Systems").Visible = True
Sheets("Total Vehicle").Visible = True

Set Gantt = Sheets("Schedule Planning")
Set sourceSheet = Sheets("TR Data")
Set cottonSheet = Sheets("Cotton Picker Specific")
Set balerSheet = Sheets("Baler Tests")
Set engineSheet = Sheets("Engine Tests")
Set cabSheet = Sheets("Cab Tests")
Set chasisSheet = Sheets("Chasis Tests")
Set powerTrainSheet = Sheets("Power Train Tests")
Set elctrcSheet = Sheets("Electrical Tests")
Set hydSheet = Sheets("Hydraulic Tests")
Set steerSheet = Sheets("Steering Systems")
Set TotlVhcl = Sheets("Total Vehicle")



'clean the slate
Gantt.Activate
With Gantt.ListObjects("PlanTable")
    Set rList = .Range
    .Unlist                           ' convert the table back to a range
End With
rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
Cells(2, 2).Value = Date


' fill Gantt

Gantt.Activate
Range(Cells(7, 1), Cells(2000, 2000)).Clear
Range(Cells(7, 15), Cells(2000, 2000)).ClearFormats

rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data
Gantt.Range(Cells(7, 1), Cells(rowCnt + 1, 8)).Clear



'Begin moving data
'rowCnt = rowCnt + 1
Gantt.Cells(7, 1) = CopyBaler()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of baler

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyCtnSpfc()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyCab()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyEngine()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyChasis()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyPwrTrn()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyElec()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyHyd()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopySteer()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyTtlVhcl()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyLabs()






' create our gantt chart
Gantt.Activate
rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row 'find the end of Data
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data
 'MsgBox (rowCnt)
'Range(Cells(8, 8), Cells(rowCnt, 9)).NumberFormat = "d-mmm-yy"
Range(Cells(7, 11), Cells(rowCnt, 667)).ClearFormats

'Convert to table
growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(6, 1), Cells(growCnt, 9)).Select


Gantt.Activate
Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.Name = "PlanTable"
With objTable
    objTable.TableStyle = "TableStyleMedium15"
End With




Range(Cells(7, 3), Cells(growCnt, 4)).ClearFormats
Range(Cells(7, 3), Cells(growCnt, 4)).NumberFormat = "d-mmm-yy"


'Sort on Date
'Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Clear
    'Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        'SortFields.Add2 Key:=Range("PlanTable[[#All],[Field Activities]]"), SortOn _
        ':=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        'SortFields.Add2 Key:=Range("PlanTable[[#All],[Scheduled Start]]"), SortOn:= _
        'xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'With ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable"). _
       'Sort
        '.Header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
        '.Apply
    'End With


'Fill out the Gantt Calender with coded colors to make it easier to find issues

Gantt.Activate
Set GanttCal = Gantt.Range(Cells(8, 16), Cells(rowCnt, colCnt))


'part is required for color coding that is in the Gantt Chart
prog = "In Progress"
strt = "To Be Started"
SPS = "Awaiting SPS Approval"
creater = "Awaiting Creator Approval"
PLPV = "Awaiting PV Approval"



'test counters
progy23 = 0
starty23 = 0
apprvy23 = 0
progy24 = 0
starty24 = 0
apprvy24 = 0


'make sure the Gannt is populated and color coded



For i = 7 To rowCnt
    Cells(i, 6).Value = "=WEEKNUM(RC[-3])"
    Cells(i, 7).Value = "=WEEKNUM(RC[-3])"
    If Cells(i, 9).Value <> "Closed" Then

    If Cells(i, 9).Value = prog Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7,RC9=""In Progress"")" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(51, 204, 204)
        End With
    End If
    
    If Cells(i, 9).Value = strt Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7,RC9=""To Be Started"")" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
        End With
    End If
    
    If Len(Cells(i, 9)) = 0 Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)
        End With
    End If
    
    If Cells(i, 9).Value = SPS Or Cells(i, 9).Value = creater Or Cells(i, 9).Value = PLPV Then
        With Gantt.Range(Cells(i, 1), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 153, 0)
        End With
    End If
       
    If Cells(i, 9).Value = "Completed" Or Cells(i, 5).Value = "Awaiting Report Approval" Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
        End With
    End If
    End If
    
 Next

  Range(Cells(7, 9), Cells(rowCnt, 9)).ClearFormats
    



'End for loop
     
growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(7, 3), Cells(growCnt, 4)).ClearFormats
Range(Cells(7, 3), Cells(growCnt, 4)).NumberFormat = "d-mmm-yy"
'Format as table
'growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
'Range(Cells(7, 14), Cells(growCnt, 15)).ClearFormats
'Range(Cells(7, 14), Cells(growCnt, 15)).NumberFormat = "d-mmm-yy"
Cells(growCnt + 3, 2) = growCnt - 14   'total test count for pending tests







'Hide tabs
Sheets("Power Train Tests").Visible = False
Sheets("Chasis Tests").Visible = False
Sheets("Baler Tests").Visible = False
Sheets("Engine Tests").Visible = False
Sheets("Cotton Picker Specific").Visible = False
Sheets("Cab Tests").Visible = False
Sheets("Electrical Tests").Visible = False
Sheets("Hydraulic Tests").Visible = False
Sheets("Steering Systems").Visible = False
Sheets("Total Vehicle").Visible = False


End Sub

