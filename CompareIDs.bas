Attribute VB_Name = "CompareIDs"
Sub CompareIDs()
Dim importID
Dim found As Range
PlanTab = "Field 2025 priority"
DataTab = "Data"
Set PlanLoc = ThisWorkbook.Worksheets(PlanTab)
Set DataLoc = ThisWorkbook.Worksheets(DataTab)
Set Plan = ThisWorkbook.Worksheets(PlanTab).ListObjects("Plan")

Dim addedRow As ListRow

'Worksheets("Data").Activate
data_rcnt = Worksheets(DataTab).Cells(Rows.Count, 1).End(xlUp).Row
plan_rcnt = Worksheets(PlanTab).Cells(Rows.Count, 1).End(xlUp).Row

'read the IDs of data location and compare them to the plan location
For i = 2 To data_rcnt
           
            If i = 91 Then
                ID = cell.Value  'this reads the CTS ID from the row coming from the data
                dummyV = yes
            Else
            End If
    
    importID = Sheets(DataTab).Cells(i, 1).Value
    PlanLoc.Columns(1).Select
    
    Set cell = Selection.Find(What:=importID, After:=Cells(i, 1), LookIn:= _
    xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)

    If cell Is Nothing Then
        plan_rcnt = Worksheets(PlanTab).Cells(Rows.Count, 1).End(xlUp).Row
    
        Set addedRow = Plan.ListRows.Add()
            For Z = 1 To 12
            If Z <> 11 Then
            With addedRow
                .Range(Z) = Sheets("Data").Cells(i, Z).Value
            End With
            Else
            End If
            
            Next
            
    Else  'fill out the table and update values as needed
        ID = cell.Value  'this reads the CTS ID from the row coming from the data
            If ID = CPTEAA095686 Then
                dummyV = yes
            Else
            End If
        For j = 4 To 12
        If j <> 11 Then
            Set found = Range("A:A").Find(ID)       ' This is looking through the plan for the CTS ID in theplan IDs

            rval = found.Row        ' gives me the row number
            Sheets(PlanTab).Cells(rval, j).Value = Sheets(DataTab).Cells(i, j).Value
        ElseIf j = 11 Then
        'Now let's set the priority based on our criteria
            Set found = Range("A:A").Find(ID)       ' This is looking through the plan for the CTS ID in theplan IDs
            rval = found.Row        ' gives me the row number
            If Cells(rval, 12).Value = "YES" And Cells(rval, 13).Value = "Yes" Then
                Sheets(PlanTab).Cells(rval, 11).Value = "High"
            ElseIf Cells(rval, 12).Value = "NO" And Cells(rval, 13).Value = "No" Then
                Sheets(PlanTab).Cells(rval, 11).Value = "Low"
            ElseIf Cells(rval, 13).Value = "" Then
                
            Else
                Sheets(PlanTab).Cells(rval, 11).Value = "Medium"
            End If
        End If
        

        Next
    End If
      
           Next
' Look at the data ID



'Next
plan_rcnt = Worksheets(PlanTab).Cells(Rows.Count, 1).End(xlUp).Row
Worksheets(PlanTab).Activate
Range(Cells(2, 5), Cells(plan_rcnt, 8)).NumberFormat = "d-mmm-yy"
End Sub

