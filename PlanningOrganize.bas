Attribute VB_Name = "PlanningOrganize"
Sub PGByDate()

Dim Gantt As Worksheet
Dim GanttCal As Range
Set Gantt = Sheets("Schedule Planning")

'Sort on Date
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Add2 Key:=Range("PlanTable[[#All],[Scheduled Start]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row 'find the end of Data
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data

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
    Cells(i, 6).Value = "=WEEKNUM(RC[2])"
    Cells(i, 7).Value = "=WEEKNUM(RC[2])"

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
       
    If Cells(i, 9).Value = "Completed" Or Cells(i, 9).Value = "Awaiting Report Approval" Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
        End With
    End If
Next
End Sub


Sub PGByName()

Dim Gantt As Worksheet
Dim GanttCal As Range
Set Gantt = Sheets("Schedule Planning")

'Sort on Date
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Add2 Key:=Range("PlanTable[[#All],[Field Activities]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row 'find the end of Data
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data

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
    Cells(i, 6).Value = "=WEEKNUM(RC[2])"
    Cells(i, 7).Value = "=WEEKNUM(RC[2])"

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
       
    If Cells(i, 9).Value = "Completed" Or Cells(i, 9).Value = "Awaiting Report Approval" Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
        End With
    End If
Next
End Sub

Sub PGByID()

Dim Gantt As Worksheet
Dim GanttCal As Range
Set Gantt = Sheets("Schedule Planning")

'Sort on ID
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable").Sort. _
        SortFields.Add2 Key:=Range("PlanTable[[#All],[TR ID'#]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Schedule Planning").ListObjects("PlanTable"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


rowCnt = Gantt.Cells(Rows.Count, 1).End(xlUp).Row 'find the end of Data
colCnt = Gantt.Cells(6, Columns.Count).End(xlToLeft).Column 'find the end of Data

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

    If Cells(i, 9).Value = prog Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7,RC5=""In Progress"")" 'RC16:R220C288
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
       
    If Cells(i, 9).Value = "Completed" Or Cells(i, 9).Value = "Awaiting Report Approval" Then
        With Gantt.Range(Cells(i, 11), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R6C>=RC6,R6C<=RC7)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
        End With
    End If
Next
End Sub


