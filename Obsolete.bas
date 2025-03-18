Attribute VB_Name = "Obsolete"
Sub RemoveObsolete()
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

'read the IDs of Plan location and compare them to the Data location
For i = 2 To plan_rcnt
    
    F_ID = Cells(i, 1).Address
    L_ID = Cells(i, 14).Address
    cellrow = i
    importID = Sheets(PlanTab).Cells(i, 1).Value        ' Set the ID to be searched for
    DataLoc.Select
    DataLoc.Columns(1).Select                           ' Where am I searching
    
    Set cell = Selection.Find(What:=importID, After:=Cells(i, 1), LookIn:= _
    xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)       ' This is the search

    If cell Is Nothing Then                 'What to do
     PlanLoc.Select
        Range(F_ID, L_ID).Select                  'Select the row for highlight
      Selection.Style = "Bad"
        'Cells(i, 1).Select
        'Plan.ListRows(i - 1).Delete     'table row is 1 less than thesheet number
        'i = i - 1       'we need to check the next cell which now holds the same address
    End If
      
           Next

End Sub



