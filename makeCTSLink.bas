Attribute VB_Name = "makeCTSLink"
Sub makeLink()
rcnt = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
Set wksh = ThisWorkbook.ActiveSheet
 
 
 
For i = 2 To rcnt
    ID = Cells(i, 5).Value
    link = "https://cts.cnhind.com:3000/#/teamcenter.search.search?searchCriteria=" & ID & "&secondaryCriteria=*"
    With wksh
        ActiveSheet.Hyperlinks.Add Anchor:=.Cells(i, 1), Address:=link, _
        TextToDisplay:=.Cells(i, 1).Value2
        End With
        
    Next
    
    
    
End Sub
