Attribute VB_Name = "PopulateGanttChart"


Sub PopulateGantt()
Attribute PopulateGantt.VB_ProcData.VB_Invoke_Func = " \n14"
'declarations
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

Set Gantt = Sheets("2024 planning")
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
Set LabTest = Sheets("Lab Tests")

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


'clean the slate
Gantt.Activate
With Gantt.ListObjects("GanttTable")
    Set rList = .Range
    .Unlist                           ' convert the table back to a range
End With
rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row

Gantt.Activate
Range(Cells(7, 1), Cells(2000, 2000)).Clear
Range(Cells(7, 15), Cells(2000, 2000)).ClearFormats

Gantt.Range(Cells(8, 1), Cells(rowCnt + 1, 14)).Clear

Cells(2, 2).Value = Date



'Begin moving data
Gantt.Cells(7, 1) = CopyBaler2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of baler

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyCtnSpfc2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row  'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyCab2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyEngine2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyChasis2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyPwrTrn2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyElec2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyHyd2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopySteer2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyTtlVhcl2024()
rowCnt = Gantt.Cells(Rows.Count, 2).End(xlUp).Row 'find the end of Data

rowCnt = rowCnt + 1
Gantt.Cells(rowCnt, 1) = CopyLabs2024()

' create our gantt chart
rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row 'find the end of Data
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data
 'MsgBox (rowCnt)
Range(Cells(8, 14), Cells(rowCnt, 15)).NumberFormat = "d-mmm-yy"
Range(Cells(8, 16), Cells(rowCnt, 667)).ClearFormats

Gantt.Activate
'Fill out the Gantt Calender with coded colors to make it easier to find issues
Set GanttCal = Gantt.Range(Cells(8, 16), Cells(rowCnt, colCnt))

prog = "In Progress"
strt = "To Be Started"
SPS = "Awaiting SPS Approval"
creater = "Awaiting Creator Approval"

'test counters
progy23 = 0
starty23 = 0
apprvy23 = 0
progy24 = 0
starty24 = 0
apprvy24 = 0

Fill = PopGantt(rowCnt, colCnt)


' fill Gantt
'For i = 8 To rowCnt
  '  If Cells(i, 12).Value = prog Then
      '  With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
      '      .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=RC14,R5C<=RC15,RC12=""In Progress"")" 'RC16:R220C288
      '      .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(51, 204, 204)
      '  End With
  '  End If
    
  '  If Cells(i, 12).Value = strt Then
    '    With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
    '        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=RC14,R5C<=RC15,RC12=""To Be Started"")" 'RC16:R220C288
     '       .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
      '  End With
    'End If
    
    'If Len(Cells(i, 12)) = 0 Then
     '   With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
     '       .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=RC14,R5C<=RC15)" 'RC16:R220C288
     '       .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)
     '   End With
   ' End If
    
   ' If Cells(i, 12).Value = SPS Or Cells(i, 12).Value = creater Then
   '     With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
    '        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=RC14,R5C<=RC15)" 'RC16:R220C288
    '        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 153, 0)
    '    End With
   ' End If
       
    'If Cells(i, 12).Value = "Completed" Or Cells(i, 12).Value = "Awaiting Report Approval" Then
   '     With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
   '         .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=RC14,R5C<=RC15)" 'RC16:R220C288
    '        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
   '     End With
   ' End If
    
    
'classify the tests by date
   ' If Cells(i, 14).Value >= "1/1/2023  1:00:00 AM" And Cells(i, 14).Value <= "12/31/2024  11:59:00 PM" Then
    '    If Cells(i, 12).Value = prog Then
     '       progy23 = progy23 + 1
   '     End If
    '    If Cells(i, 12).Value = strt Then
    '        starty23 = starty23 + 1
    '    End If
     '   If Cells(i, 12).Value = SPS Or Cells(i, 12).Value = creater Then
     '       apprvy23 = apprvy23 + 1
     '   End If
   ' End If
   ' If Cells(i, 14).Value >= "1/1/2024  1:00:00 AM" And Cells(i, 14).Value <= "12/31/2024  11:59:00 PM" Then
    '    If Cells(i, 12).Value = prog Then
    '        progy24 = progy24 + 1
     '   End If
     '   If Cells(i, 12).Value = strt Then
     '       starty24 = starty24 + 1
      '  End If
      '  If Cells(i, 12).Value = SPS Or Cells(i, 12).Value = creater Then
      '      apprvy24 = apprvy24 + 1
       ' End If
    'End If
'Next
    
    
    
'format dates into desired format
growCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(7, 14), Cells(growCnt, 15)).ClearFormats
Range(Cells(7, 14), Cells(growCnt, 15)).NumberFormat = "d-mmm-yy"
Cells(growCnt + 3, 2) = growCnt - 16    'total test count for pending tests



Range(Cells(6, 1), Cells(growCnt, 15)).Select


Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.Name = "GanttTable"
With objTable
    objTable.TableStyle = "TableStyleMedium15"
End With


' highlight late tests
late = TestAlert2024()


Cells(growCnt + 4, 1) = "In Progress"   'total test count for pending tests
Cells(growCnt + 5, 1) = "To Be Started"    'total test count for pending tests
Cells(growCnt + 6, 1) = "Waiting for approval"    'total test count for pending tests
Cells(growCnt + 4, 2) = progy23    'total test count for pending tests
Cells(growCnt + 5, 2) = starty23    'total test count for pending tests
Cells(growCnt + 6, 2) = apprvy23    'total test count for pending tests
Cells(growCnt + 4, 3) = progy24    'total test count for pending tests
Cells(growCnt + 5, 3) = starty24    'total test count for pending tests
Cells(growCnt + 6, 3) = apprvy24    'total test count for pending tests

Cells(growCnt + 7, 2) = Application.Sum(Range(Cells(growCnt + 4, 2), Cells(growCnt + 6, 2)))
'total test count for pending tests
Cells(growCnt + 7, 3) = Application.Sum(Range(Cells(growCnt + 4, 3), Cells(growCnt + 6, 3))) 'total test count for pending tests"
Cells(growCnt + 7, 14) = Cells(growCnt + 7, 2).Value + Cells(growCnt + 7, 3) 'total test count for pending tests

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

'Sheets("Test OverView").Visible = True

End Sub
