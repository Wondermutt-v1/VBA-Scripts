Attribute VB_Name = "PopGanttFxn"
Function PopGantt(rowCnt, colCnt)

Set Gantt = Sheets("2024 planning")

statusCol = 12
startDateCol = 14
endDateCol = 15
startDate = RC14
endDate = RC15

prog = "In Progress"
strt = "To Be Started"
SPS = "Awaiting SPS Approval"
creater = "Awaiting Creator Approval"


i = 0
' fill Gantt
For i = 8 To rowCnt
    If Cells(i, statusCol).Value = prog Then
        With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=startDate,R5C<=endDate,RC12=""In Progress"")" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(51, 204, 204)
        End With
    End If
    
    If Cells(i, 12).Value = strt Then
        With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=startDate,R5C<=endDate,RC12=""To Be Started"")" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
        End With
    End If
    
    If Len(Cells(i, 12)) = 0 Then
        With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=startDate,R5C<=endDate)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)
        End With
    End If
    
    If Cells(i, 12).Value = SPS Or Cells(i, statusCol).Value = creater Then
        With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=startDate,R5C<=endDate)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 153, 0)
        End With
    End If
       
    If Cells(i, 12).Value = "Completed" Or Cells(i, statusCol).Value = "Awaiting Report Approval" Then
        With Gantt.Range(Cells(i, 17), Cells(i, colCnt))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(R5C>=startDate,R5C<=endDate)" 'RC16:R220C288
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(18, 228, 128)
        End With
    End If
    
    
'classify the tests by date
    If Cells(i, startDateCol).Value >= "1/1/2023  1:00:00 AM" And Cells(i, 14).Value <= "12/31/2024  11:59:00 PM" Then
        If Cells(i, statusCol).Value = prog Then
            progy23 = progy23 + 1
        End If
        If Cells(i, statusCol).Value = strt Then
            starty23 = starty23 + 1
        End If
        If Cells(i, statusCol).Value = SPS Or Cells(i, statusCol).Value = creater Then
            apprvy23 = apprvy23 + 1
        End If
    End If
    If Cells(i, 14).Value >= "1/1/2024  1:00:00 AM" And Cells(i, startDateCol).Value <= "12/31/2024  11:59:00 PM" Then
        If Cells(i, statusCol).Value = prog Then
            progy24 = progy24 + 1
        End If
        If Cells(i, statusCol).Value = strt Then
            starty24 = starty24 + 1
        End If
        If Cells(i, statusCol).Value = SPS Or Cells(i, statusCol).Value = creater Then
            apprvy24 = apprvy24 + 1
        End If
    End If
Next

End Function


