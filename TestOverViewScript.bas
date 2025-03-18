Attribute VB_Name = "TestOverViewScript"
Sub TestOverview()
Dim WB As Workbook
Dim SrcSht As Worksheet
Dim DestSht As Worksheet
Dim Bucket() As Integer


Set SrcSht = ThisWorkbook.Worksheets("TR Data")
'Dim DestSht As Worksheet

Set WB = ThisWorkbook
Set DestSht = ThisWorkbook.Worksheets("Test Overview")
DestSht.Activate
rowCnt = SrcSht.Cells(Rows.Count, 2).End(xlUp).Row

prog = "In Progress"
strt = "To Be Started"
SPS = "Awaiting SPS Approval"
creater = "Awaiting Creator Approval"
comp = "Completed"
cls = "Closed"
j = 5

y23 = 0
y24 = 0
y25 = 0

'month
Jan23 = 0
progJan23 = 0
strtJan23 = 0
compJan23 = 0
pubJan23 = 0
Feb23 = 0
progFeb23 = 0
strtFeb23 = 0
compFeb23 = 0
pubFeb23 = 0
Mar23 = 0
progMar23 = 0
strtMar23 = 0
compMar23 = 0
pubMar23 = 0
Apr23 = 0
progApr23 = 0
strtApr23 = 0
compApr23 = 0
pubApr23 = 0
May23 = 0
progMay23 = 0
strtMay23 = 0
compMay23 = 0
pubMay23 = 0
Jun23 = 0
progJun23 = 0
strtJun23 = 0
compJun23 = 0
pubJun23 = 0
Jul23 = 0
progJul23 = 0
strtJul23 = 0
compJul23 = 0
pubJul23 = 0
Aug23 = 0
progAug23 = 0
strtAug23 = 0
compAug23 = 0
pubAug23 = 0
Sep23 = 0
progSep23 = 0
strtSep23 = 0
compSep23 = 0
pubSep23 = 0
Oct23 = 0
progOct23 = 0
strtOct23 = 0
compOct23 = 0
pubOct23 = 0
Nov23 = 0
progNov23 = 0
strtNov23 = 0
compNov23 = 0
pubNov23 = 0
Dec23 = 0
progDec23 = 0
strtDec23 = 0
compDec23 = 0
pubDec23 = 0
Jan24 = 0
progJan24 = 0
strtJan24 = 0
compJan24 = 0
pubJan24 = 0
Feb24 = 0
progFeb24 = 0
strtFeb24 = 0
compFeb24 = 0
pubFeb24 = 0
Mar24 = 0
progMar24 = 0
strtMar24 = 0
compMar24 = 0
pubMar24 = 0
Apr24 = 0
progApr24 = 0
strtApr24 = 0
compApr24 = 0
pubApr24 = 0
May24 = 0
progMay24 = 0
strtMay24 = 0
compMay24 = 0
pubMay24 = 0
Jun24 = 0
progJun24 = 0
strtJun24 = 0
compJun24 = 0
pubJun24 = 0
Jul24 = 0
progJul24 = 0
strtJul24 = 0
compJul24 = 0
pubJul24 = 0
Aug24 = 0
progAug24 = 0
strtAug24 = 0
compAug24 = 0
pubAug24 = 0
Sep24 = 0
progSep24 = 0
strtSep24 = 0
compSep24 = 0
pubSep24 = 0
Oct24 = 0
progOct24 = 0
strtOct24 = 0
compOct24 = 0
pubOct24 = 0
Nov24 = 0
progNov24 = 0
strtNov24 = 0
compNov24 = 0
pubNov24 = 0
Dec24 = 0
progDec24 = 0
strtDec24 = 0
compDec24 = 0
pubDec24 = 0
Jan25 = 0
progJan25 = 0
strtJan25 = 0
compJan25 = 0
pubJan25 = 0
Feb25 = 0
progFeb25 = 0
strtFeb25 = 0
compFeb25 = 0
pubFeb25 = 0
Mar25 = 0
progMar25 = 0
strtMar25 = 0
compMar25 = 0
pubMar25 = 0
Apr25 = 0
progApr25 = 0
strtApr25 = 0
compApr25 = 0
pubApr25 = 0
May25 = 0
progMay25 = 0
strtMay25 = 0
compMay25 = 0
pubMay25 = 0
Jun25 = 0
progJun25 = 0
strtJun25 = 0
compJun25 = 0
pubJun25 = 0
Jul25 = 0
progJul25 = 0
strtJul25 = 0
compJul25 = 0
pubJul25 = 0
Aug25 = 0
progAug25 = 0
strtAug25 = 0
compAug25 = 0
pubAug25 = 0
Sep25 = 0
progSep25 = 0
strtSep25 = 0
compSep25 = 0
pubSep25 = 0
Oct25 = 0
progOct25 = 0
strtOct25 = 0
compOct25 = 0
pubOct25 = 0
Nov25 = 0
progNov25 = 0
strtNov25 = 0
compNov25 = 0
pubNov25 = 0
Dec25 = 0
progDec25 = 0
strtDec25 = 0
compDec25 = 0
pubDec25 = 0

Untest = 0
'DestSht = Sheets("Test Overview")
For i = 5 To rowCnt
Cells(43, 1) = i
    If SrcSht.Cells(i, 7) = "No Longer Required" Then
    Untest = Untest + 1
    End If
    If SrcSht.Cells(i, 7) <> "No Longer Required" Then
        'DestSht.Cells(j, 1) = WorksheetFunction.WeekNum(SrcSht.Cells(i, 20), 1)
        'DestSht.Cells(j, 2) = SrcSht.Cells(i, 20)
        'DestSht.Cells(j, 2).NumberFormat = "yy"
        'Set wk = DestSht.Cells(j, 1)
        'Set Yr = SrcSht.Cells(i, 20).Value
        mth = SrcSht.Cells(i, 20).Value
        mont = Month(mth)
        
        Yr = Year(mth)
        'DestSht.Cells(j, 3).Value = Yr & "-" & wk 'DestSht.Cells(j, 2)
        'DestSht.Cells(j, 4).Value = SrcSht.Cells(i, 6)
        
       
    
        If Yr = 2023 Then
            If Month(SrcSht.Cells(i, 20).Value) = 1 Then
                  
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJan23 = progJan23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJan23 = strtJan23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJan23 = compJan23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJan23 = pubJan23 + 1
                    End If
    
                Jan23 = Jan23 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 2 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progFeb23 = progFeb23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtFeb23 = strtFeb23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compFeb23 = compFeb23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubFeb23 = pubFeb23 + 1
                    End If
    
                Feb23 = Feb23 + 1
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 3 Then
            
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMar23 = progMar23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMar23 = strtMar23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMar23 = compMar23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMar23 = pubMar23 + 1
                    End If
                Mar23 = Mar23 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 4 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progApr23 = progApr23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtApr23 = strtApr23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compApr23 = compApr23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubApr23 = pubApr23 + 1
                    End If
                Apr23 = Apr23 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 5 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMay23 = progMay23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMay23 = strtMay23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMay23 = compMay23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMay23 = pubMay23 + 1
                    End If
                May23 = May23 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 6 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJun23 = progJun23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJun23 = strtJun23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJun23 = compJun23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJun23 = pubJun23 + 1
                    End If
                Jun23 = Jun23 + 1
            End If
            If Month(SrcSht.Cells(i, 20).Value) = 7 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJul23 = progJul23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJul23 = strtJul23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJul23 = compJul23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJul23 = pubJul23 + 1
                    End If
                Jul23 = Jul23 + 1
            End If
           '
            If Month(SrcSht.Cells(i, 20).Value) = 8 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progAug23 = progAug23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtAug23 = strtAug23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compAug23 = compAug23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubAug23 = pubAug23 + 1
                    End If
                Aug23 = Aug23 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 9 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progSep23 = progSep23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtSep23 = strtSep23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compSep23 = compSep23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubSep23 = pubSep23 + 1
                    End If
                Sep23 = Sep23 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 10 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progOct23 = progOct23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtOct23 = strtOct23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compOct23 = compOct23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubOct23 = pubOct23 + 1
                    End If
                Oct23 = Oct23 + 1
                
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 11 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progNov23 = progNov23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtNov23 = strtNov23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compNov23 = compNov23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubNov23 = pubNov23 + 1
                    End If
                Nov23 = Nov23 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 12 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progDec23 = progDec23 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtDec23 = strtDec23 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compDec23 = compDec23 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubDec23 = pubDec23 + 1
                    End If
                Dec23 = Dec23 + 1
            End If
            y23 = y23 + 1
            j = j + 1
        'For i = 8 To rowCnt
        End If
        'End If
        
        If Year(SrcSht.Cells(i, 20).Value) = 2024 Then
            If Month(SrcSht.Cells(i, 20).Value) = 1 Then
                  
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJan24 = progJan24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJan24 = strtJan24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJan24 = compJan24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJan24 = pubJan24 + 1
                    End If
    
                Jan24 = Jan24 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 2 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progFeb24 = progFeb24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtFeb24 = strtFeb24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compFeb24 = compFeb24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubFeb24 = pubFeb24 + 1
                    End If
    
                Feb24 = Feb24 + 1
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 3 Then
            
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMar24 = progMar24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMar24 = strtMar24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMar24 = compMar24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMar24 = pubMar24 + 1
                    End If
                Mar24 = Mar24 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 4 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progApr24 = progApr24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtApr24 = strtApr24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compApr24 = compApr24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubApr24 = pubApr24 + 1
                    End If
                Apr24 = Apr24 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 5 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMay24 = progMay24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMay24 = strtMay24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMay24 = compMay24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMay24 = pubMay24 + 1
                    End If
                May24 = May24 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 6 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJun24 = progJun24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJun24 = strtJun24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJun24 = compJun24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJun24 = pubJun24 + 1
                    End If
                Jun24 = Jun24 + 1
            End If
            If Month(SrcSht.Cells(i, 20).Value) = 7 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJul24 = progJul24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJul24 = strtJul24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJul24 = compJul24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJul24 = pubJul24 + 1
                    End If
                Jul24 = Jul24 + 1
            End If
           '
            If Month(SrcSht.Cells(i, 20).Value) = 8 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progAug24 = progAug24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtAug24 = strtAug24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compAug24 = compAug24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubAug24 = pubAug24 + 1
                    End If
                Aug24 = Aug24 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 9 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progSep24 = progSep24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtSep24 = strtSep24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compSep24 = compSep24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubSep24 = pubSep24 + 1
                    End If
                Sep24 = Sep24 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 10 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progOct24 = progOct24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtOct24 = strtOct24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compOct24 = compOct24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubOct24 = pubOct24 + 1
                    End If
                Oct24 = Oct24 + 1
                
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 11 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progNov24 = progNov24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtNov24 = strtNov24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compNov24 = compNov24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubNov24 = pubNov24 + 1
                    End If
                Nov24 = Nov24 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 12 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progDec24 = progDec24 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtDec24 = strtDec24 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compDec24 = compDec24 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubDec24 = pubDec24 + 1
                    End If
                Dec24 = Dec24 + 1
            End If
            'End If
            y24 = y24 + 1
        End If
        'End If
        
         If Year(SrcSht.Cells(i, 20).Value) = 2025 Then
            If Month(SrcSht.Cells(i, 20).Value) = 1 Then
                  
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJan25 = progJan25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJan25 = strtJan25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJan25 = compJan25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJan25 = pubJan25 + 1
                    End If
    
                Jan25 = Jan25 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 2 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progFeb25 = progFeb25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtFeb25 = strtFeb25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compFeb25 = compFeb25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubFeb25 = pubFeb25 + 1
                    End If
    
                Feb25 = Feb25 + 1
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 3 Then
            
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMar25 = progMar25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMar25 = strtMar25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMar25 = compMar25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMar25 = pubMar25 + 1
                    End If
                Mar25 = Mar25 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 4 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progApr25 = progApr25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtApr25 = strtApr25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compApr25 = compApr25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubApr25 = pubApr25 + 1
                    End If
                Apr25 = Apr25 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 5 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progMay25 = progMay25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtMay25 = strtMay25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compMay25 = compMay25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubMay25 = pubMay25 + 1
                    End If
                May25 = May25 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 6 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJun25 = progJun25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJun25 = strtJun25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJun25 = compJun25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJun25 = pubJun25 + 1
                    End If
                Jun25 = Jun25 + 1
            End If
            If Month(SrcSht.Cells(i, 20).Value) = 7 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progJul25 = progJul25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtJul25 = strtJul25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compJul25 = compJul25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubJul25 = pubJul25 + 1
                    End If
                Jul25 = Jul25 + 1
            End If
           '
            If Month(SrcSht.Cells(i, 20).Value) = 8 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progAug25 = progAug25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtAug25 = strtAug25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compAug25 = compAug25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubAug25 = pubAug25 + 1
                    End If
                Aug25 = Aug25 + 1
            End If
            '
            If Month(SrcSht.Cells(i, 20).Value) = 9 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progSep25 = progSep25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtSep25 = strtSep25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compSep25 = compSep25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubSep25 = pubSep25 + 1
                    End If
                Sep25 = Sep25 + 1
            End If
        
            If Month(SrcSht.Cells(i, 20).Value) = 10 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progOct25 = progOct25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtOct25 = strtOct25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compOct25 = compOct25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubOct25 = pubOct25 + 1
                    End If
                Oct25 = Oct25 + 1
                
            End If
            
            If Month(SrcSht.Cells(i, 20).Value) = 11 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progNov25 = progNov25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtNov25 = strtNov25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compNov25 = compNov25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubNov25 = pubNov25 + 1
                    End If
                Nov25 = Nov25 + 1
            End If
            
             If Month(SrcSht.Cells(i, 20).Value) = 12 Then
                    If SrcSht.Cells(i, 7).Value = prog Then
                        progDec25 = progDec25 + 1
                    End If
    
                    If SrcSht.Cells(i, 7).Value = strt Then
                        strtDec25 = strtDec25 + 1
                    End If
                    If SrcSht.Cells(i, 7).Value = comp Or SrcSht.Cells(i, 7).Value = "Closed" Then
                        compDec25 = compDec25 + 1
                    End If
                    If Len(SrcSht.Cells(i, 7)) = 0 Or SrcSht.Cells(i, 7).Value = SPS Or SrcSht.Cells(i, 7).Value = creater Then
                        pubDec25 = pubDec25 + 1
                    End If
                Dec25 = Dec25 + 1
            End If
            'End If
            y25 = y25 + 1
        End If
   End If
Next

'Data table


'23
DestSht.Cells(5, 9).Value = "1-Jan-23"
DestSht.Cells(5, 10).Value = Jan23
DestSht.Cells(5, 11).Value = strtJan23
DestSht.Cells(5, 12).Value = progJan23
DestSht.Cells(5, 13).Value = compJan23
DestSht.Cells(5, 14).Value = pubJan23

DestSht.Cells(6, 9).Value = "1-Feb-23"
DestSht.Cells(6, 10).Value = Feb23
DestSht.Cells(6, 11).Value = strtFeb23
DestSht.Cells(6, 12).Value = progFeb23
DestSht.Cells(6, 13).Value = compFeb23
DestSht.Cells(6, 14).Value = pubFeb23

DestSht.Cells(7, 9).Value = "1-Mar-23"
DestSht.Cells(7, 10).Value = Mar23
DestSht.Cells(7, 11).Value = strtMar23
DestSht.Cells(7, 12).Value = progMar23
DestSht.Cells(7, 13).Value = compMar23
DestSht.Cells(7, 14).Value = pubMar23

DestSht.Cells(8, 9).Value = "1-Apr-23"
DestSht.Cells(8, 10).Value = Apr23
DestSht.Cells(8, 11).Value = strtApr23
DestSht.Cells(8, 12).Value = progApr23
DestSht.Cells(8, 13).Value = compApr23
DestSht.Cells(8, 14).Value = pubApr23

DestSht.Cells(9, 9).Value = "1-May-23"
DestSht.Cells(9, 10).Value = May23
DestSht.Cells(9, 11).Value = strtMay23
DestSht.Cells(9, 12).Value = progMay23
DestSht.Cells(9, 13).Value = compMay23
DestSht.Cells(9, 14).Value = pubMay23

DestSht.Cells(10, 9).Value = "1-Jun-23"
DestSht.Cells(10, 10).Value = Jun23
DestSht.Cells(10, 11).Value = strtJun23
DestSht.Cells(10, 12).Value = progJun23
DestSht.Cells(10, 13).Value = compJun23
DestSht.Cells(10, 14).Value = pubJun23

DestSht.Cells(11, 9).Value = "1-Jul-23"
DestSht.Cells(11, 10).Value = Jul23
DestSht.Cells(11, 11).Value = strtJul23
DestSht.Cells(11, 12).Value = progJul23
DestSht.Cells(11, 13).Value = compJul23
DestSht.Cells(11, 14).Value = pubJul23

DestSht.Cells(12, 9).Value = "1-Aug-23"
DestSht.Cells(12, 10).Value = Aug23
DestSht.Cells(12, 11).Value = strtAug23
DestSht.Cells(12, 12).Value = progAug23
DestSht.Cells(12, 13).Value = compAug23
DestSht.Cells(12, 14).Value = pubAug23

DestSht.Cells(13, 9).Value = "1-Sep-23"
DestSht.Cells(13, 10).Value = Sep23
DestSht.Cells(13, 11).Value = strtSep23
DestSht.Cells(13, 12).Value = progSep23
DestSht.Cells(13, 13).Value = compSep23
DestSht.Cells(13, 14).Value = pubSep23

DestSht.Cells(14, 9).Value = "1-Oct-23"
DestSht.Cells(14, 10).Value = Oct23
DestSht.Cells(14, 11).Value = strtOct23
DestSht.Cells(14, 12).Value = progOct23
DestSht.Cells(14, 13).Value = compOct23
DestSht.Cells(14, 14).Value = pubOct23

DestSht.Cells(15, 9).Value = "1-Nov-23"
DestSht.Cells(15, 10).Value = Nov23
DestSht.Cells(15, 11).Value = strtNov23
DestSht.Cells(15, 12).Value = progNov23
DestSht.Cells(15, 13).Value = compNov23
DestSht.Cells(15, 14).Value = pubNov23

DestSht.Cells(16, 9).Value = "1-Dec-23"
DestSht.Cells(16, 10).Value = Dec23
DestSht.Cells(16, 11).Value = strtDec23
DestSht.Cells(16, 12).Value = progDec23
DestSht.Cells(16, 13).Value = compDec23
DestSht.Cells(16, 14).Value = pubDec23

'24

DestSht.Cells(17, 9).Value = "1-Jan-24"
DestSht.Cells(17, 10).Value = Jan24
DestSht.Cells(17, 11).Value = strtJan24
DestSht.Cells(17, 12).Value = progJan24
DestSht.Cells(17, 13).Value = compJan24
DestSht.Cells(17, 14).Value = pubJan24

DestSht.Cells(18, 9).Value = "1-Feb-24"
DestSht.Cells(18, 10).Value = Feb24
DestSht.Cells(18, 11).Value = strtFeb24
DestSht.Cells(18, 12).Value = progFeb24
DestSht.Cells(18, 13).Value = compFeb24
DestSht.Cells(18, 14).Value = pubFeb24

DestSht.Cells(19, 9).Value = "1-Mar-24"
DestSht.Cells(19, 10).Value = Mar24
DestSht.Cells(19, 11).Value = strtMar24
DestSht.Cells(19, 12).Value = progMar24
DestSht.Cells(19, 13).Value = compMar24
DestSht.Cells(19, 14).Value = pubMar24

DestSht.Cells(20, 9) = "1-Apr-24"
DestSht.Cells(20, 10) = Apr24
DestSht.Cells(20, 11) = strtApr24
DestSht.Cells(20, 12) = progApr24
DestSht.Cells(20, 13) = compApr24
DestSht.Cells(20, 14).Value = pubApr24

DestSht.Cells(21, 9).Value = "1-May-24"
DestSht.Cells(21, 10).Value = May24
DestSht.Cells(21, 11).Value = strtMay24
DestSht.Cells(21, 12).Value = progMay24
DestSht.Cells(21, 13).Value = compMay24
DestSht.Cells(21, 14).Value = pubMay24

DestSht.Cells(22, 9).Value = "1-Jun-24"
DestSht.Cells(22, 10).Value = Jun24
DestSht.Cells(22, 11).Value = strtJun24
DestSht.Cells(22, 12).Value = progJun24
DestSht.Cells(22, 13).Value = compJun24
DestSht.Cells(22, 14).Value = pubJun24

DestSht.Cells(23, 9).Value = "1-Jul-24"
DestSht.Cells(23, 10).Value = Jul24
DestSht.Cells(23, 11).Value = strtJul24
DestSht.Cells(23, 12).Value = progJul24
DestSht.Cells(23, 13).Value = compJul24
DestSht.Cells(23, 14).Value = pubJul24


DestSht.Cells(24, 9).Value = "1-Aug-24"
DestSht.Cells(24, 10).Value = Aug24
DestSht.Cells(24, 11).Value = strtAug24
DestSht.Cells(24, 12).Value = progAug24
DestSht.Cells(24, 13).Value = compAug24
DestSht.Cells(24, 14).Value = pubAug24

DestSht.Cells(25, 9).Value = "1-Sep-24"
DestSht.Cells(25, 10).Value = Sep24
DestSht.Cells(25, 11).Value = strtSep24
DestSht.Cells(25, 12).Value = progSep24
DestSht.Cells(25, 13).Value = compSep24
DestSht.Cells(25, 14).Value = pubSep24

DestSht.Cells(26, 9).Value = "1-Oct-24"
DestSht.Cells(26, 10).Value = Oct24
DestSht.Cells(26, 11).Value = strtOct24
DestSht.Cells(26, 12).Value = progOct24
DestSht.Cells(26, 13).Value = compOct24
DestSht.Cells(26, 14).Value = pubOct24

DestSht.Cells(27, 9).Value = "1-Nov-24"
DestSht.Cells(27, 10).Value = Nov24
DestSht.Cells(27, 11).Value = strtNov24
DestSht.Cells(27, 12).Value = progNov24
DestSht.Cells(27, 13).Value = compNov24
DestSht.Cells(27, 14).Value = pubNov24

DestSht.Cells(28, 9).Value = "1-Dec-24"
DestSht.Cells(28, 10).Value = Dec24
DestSht.Cells(28, 11).Value = strtDec24
DestSht.Cells(28, 12).Value = progDec24
DestSht.Cells(28, 13).Value = compDec24
DestSht.Cells(28, 14).Value = pubDec24

'25
DestSht.Cells(29, 9).Value = "1-Jan-25"
DestSht.Cells(29, 10).Value = Jan25
DestSht.Cells(29, 11).Value = strtJan25
DestSht.Cells(29, 12).Value = progJan25
DestSht.Cells(29, 13).Value = compJan25
DestSht.Cells(29, 14).Value = pubJan25
'DestSht.Cells(29, 15).Value =

DestSht.Cells(30, 9).Value = "1-Feb-25"
DestSht.Cells(30, 10).Value = Feb25
DestSht.Cells(30, 11).Value = strtFeb25
DestSht.Cells(30, 12).Value = progFeb25
DestSht.Cells(30, 13).Value = compFeb25
DestSht.Cells(30, 14).Value = pubFeb25
'DestSht.Cells(30, 15).Value =

DestSht.Cells(31, 9).Value = "1-Mar-25"
DestSht.Cells(31, 10).Value = Mar25
DestSht.Cells(31, 11).Value = strtMar25
DestSht.Cells(31, 12).Value = progMar25
DestSht.Cells(31, 13).Value = compMar25
DestSht.Cells(31, 14).Value = pubMar25
'DestSht.Cells(31, 15).Value =

DestSht.Cells(32, 9).Value = "1-Apr-25"
DestSht.Cells(32, 10).Value = Apr25
DestSht.Cells(32, 11).Value = strtApr25
DestSht.Cells(32, 12).Value = progApr25
DestSht.Cells(32, 13).Value = compApr25
DestSht.Cells(32, 14).Value = pubApr25
'DestSht.Cells(32, 15).Value =

DestSht.Cells(33, 9).Value = "1-May-25"
DestSht.Cells(33, 10).Value = May25
DestSht.Cells(33, 11).Value = strtMay25
DestSht.Cells(33, 12).Value = progMay25
DestSht.Cells(33, 13).Value = compMay25
DestSht.Cells(33, 14).Value = pubMay25
'DestSht.Cells(33, 15).Value =

DestSht.Cells(34, 9).Value = "1-Jun-25"
DestSht.Cells(34, 10).Value = Jun25
DestSht.Cells(34, 11).Value = strtJun25
DestSht.Cells(34, 12).Value = progJun25
DestSht.Cells(34, 13).Value = compJun25
DestSht.Cells(34, 14).Value = pubJun25

DestSht.Cells(35, 9).Value = "1-Jul-25"
DestSht.Cells(35, 10).Value = Jul25
DestSht.Cells(35, 11).Value = strtJul25
DestSht.Cells(35, 12).Value = progJul25
DestSht.Cells(35, 13).Value = compJul25
DestSht.Cells(35, 14).Value = pubJul25
'DestSht.Cells(35, 15).Value =

DestSht.Cells(36, 9).Value = "1-Aug-25"
DestSht.Cells(36, 11).Value = Aug25
DestSht.Cells(36, 12).Value = strtAug25
DestSht.Cells(36, 13).Value = progAug25
DestSht.Cells(36, 14).Value = compAug25
DestSht.Cells(36, 15).Value = pubAug25

DestSht.Cells(37, 9).Value = "1-Sep-25"
DestSht.Cells(37, 10).Value = Sep25
DestSht.Cells(37, 11).Value = strtSep25
DestSht.Cells(37, 12).Value = progSep25
DestSht.Cells(37, 13).Value = compSep25
DestSht.Cells(37, 14).Value = pubSep25

DestSht.Cells(38, 9).Value = "1-Oct-25"
DestSht.Cells(38, 10).Value = Oct25
DestSht.Cells(38, 11).Value = strtOct25
DestSht.Cells(38, 12).Value = progOct25
DestSht.Cells(38, 13).Value = compOct25
DestSht.Cells(38, 14).Value = pubOct25


DestSht.Cells(39, 9).Value = "1-Nov-25"
DestSht.Cells(39, 10).Value = Nov25
DestSht.Cells(39, 11).Value = strtNov25
DestSht.Cells(39, 12).Value = progNov25
DestSht.Cells(39, 13).Value = compNov25
DestSht.Cells(39, 14).Value = pubNov25

DestSht.Cells(40, 9).Value = "1-Dec-25"
DestSht.Cells(40, 10).Value = Dec25
DestSht.Cells(40, 11).Value = strtDec25
DestSht.Cells(40, 12).Value = progDec25
DestSht.Cells(40, 13).Value = compDec25
DestSht.Cells(40, 14).Value = pubDec25


DestSht.Cells(5, 5).Value = y23
DestSht.Cells(6, 5).Value = y24
DestSht.Cells(7, 5).Value = y25
DestSht.Cells(8, 5).Value = y23 + y24 + y25


'create graph
'Sheets(“Cumulative Test Chart”).ChartTitle.Text = “Cumulative Test by Month”
'With Worksheets("Cummulative Test Cart").ChartObjects(1).Chart
 '.HasTitle = True
 '.ChartTitle.Text = "1995 Rainfall Totals by Month"
'End With

DestSht.Cells(5, 1).Value = "1-Jan-23"
DestSht.Cells(5, 2).Value = Jan23
DestSht.Cells(5, 5).Value = compJan23
DestSht.Cells(5, 4).Value = progJan23 + compJan23
DestSht.Cells(5, 3).Value = strtJan23 + progJan23 + compJan23
'DestSht.Cells(5, 15).Value = pubJan23

DestSht.Cells(6, 1).Value = "1-Feb-23"
DestSht.Cells(6, 2).Value = Feb23
DestSht.Cells(6, 5).Value = compFeb23
DestSht.Cells(6, 4).Value = progFeb23 + compFeb23
DestSht.Cells(6, 3).Value = strtFeb23 + progFeb23 + compFeb23 'DestSht.Cells(6, 15).Value = pubFeb23

DestSht.Cells(7, 1).Value = "1-Mar-23"
DestSht.Cells(7, 2).Value = Mar23
DestSht.Cells(7, 3).Value = strtMar23 + progMar23 + compMar23
DestSht.Cells(7, 4).Value = progMar23 + compMar23
DestSht.Cells(7, 5).Value = compMar23
'DestSht.Cells(7, 15).Value = pubMar23

DestSht.Cells(8, 1).Value = "1-Apr-23"
DestSht.Cells(8, 2).Value = Apr23
DestSht.Cells(8, 3).Value = strtApr23 + progApr23 + compApr23
DestSht.Cells(8, 4).Value = progApr23 + compApr23
DestSht.Cells(8, 5).Value = compApr23
'DestSht.Cells(8, 15).Value = pubApr23

DestSht.Cells(9, 1).Value = "1-May-23"
DestSht.Cells(9, 2).Value = May23
DestSht.Cells(9, 3).Value = strtMay23 + progMay23 + compMay23
DestSht.Cells(9, 4).Value = progMay23 + compMay23
DestSht.Cells(9, 5).Value = compMay23
'DestSht.Cells(9, 15).Value = pubMay23

DestSht.Cells(10, 1).Value = "1-Jun-23"
DestSht.Cells(10, 2).Value = Jun23
DestSht.Cells(10, 3).Value = strtJun23 + progJun23 + compJun23
DestSht.Cells(10, 4).Value = progJun23 + compJun23
DestSht.Cells(10, 5).Value = compJun23
'DestSht.Cells(10, 15).Value = pubJun23

DestSht.Cells(11, 1).Value = "1-Jul-23"
DestSht.Cells(11, 2).Value = Jul23
DestSht.Cells(11, 3).Value = strtJul23 + progJul23 + compJul23
DestSht.Cells(11, 4).Value = progJul23 + compJul23
DestSht.Cells(11, 5).Value = compJul23
'DestSht.Cells(11, 15).Value = pubJul23

DestSht.Cells(12, 1).Value = "1-Aug-23"
DestSht.Cells(12, 2).Value = Aug23
DestSht.Cells(12, 3).Value = strtAug23 + progAug23 + compAug23
DestSht.Cells(12, 4).Value = progAug23 + compAug23
DestSht.Cells(12, 5).Value = compAug23
'DestSht.Cells(12, 15).Value = pubAug23

DestSht.Cells(13, 1).Value = "1-Sep-23"
DestSht.Cells(13, 2).Value = Sep23
DestSht.Cells(13, 3).Value = strtSep23 + progSep23 + compSep23
DestSht.Cells(13, 4).Value = progSep23 + compSep23
DestSht.Cells(13, 5).Value = compSep23
'DestSht.Cells(13, 15).Value = pubSep23

DestSht.Cells(14, 1).Value = "1-Oct-23"
DestSht.Cells(14, 2).Value = Oct23
DestSht.Cells(14, 3).Value = strtOct23 + progOct23 + compOct23
DestSht.Cells(14, 4).Value = progOct23 + compOct23
DestSht.Cells(14, 5).Value = compOct23
'DestSht.Cells(14, 15).Value = pubOct23

DestSht.Cells(15, 1).Value = "1-Nov-23"
DestSht.Cells(15, 2).Value = Nov23
DestSht.Cells(15, 3).Value = strtNov23 + progNov23 + compNov23
DestSht.Cells(15, 4).Value = progNov23 + compNov23
DestSht.Cells(15, 5).Value = compNov23
'DestSht.Cells(15, 15).Value = pubNov23

DestSht.Cells(16, 1).Value = "1-Dec-23"
DestSht.Cells(16, 2).Value = Dec23
DestSht.Cells(16, 3).Value = strtDec23 + progDec23 + compDec23
DestSht.Cells(16, 4).Value = progDec23 + compDec23
DestSht.Cells(16, 5).Value = compDec23
'DestSht.Cells(16, 15).Value = pubDec23

DestSht.Cells(17, 1).Value = "1-Jan-24"
DestSht.Cells(17, 2).Value = Jan24
DestSht.Cells(17, 3).Value = strtJan24 + progJan24 + compJan24
DestSht.Cells(17, 4).Value = progJan24 + compJan24
DestSht.Cells(17, 5).Value = compJan24
'DestSht.Cells(5, 15).Value = pubJan23

DestSht.Cells(18, 1).Value = "1-Feb-24"
DestSht.Cells(18, 2).Value = Feb24
DestSht.Cells(18, 3).Value = strtFeb24 + progFeb24 + compFeb24
DestSht.Cells(18, 4).Value = progFeb24 + compFeb24
DestSht.Cells(18, 5).Value = compFeb24
'DestSht.Cells(6, 15).Value = pubFeb23

DestSht.Cells(19, 1).Value = "1-Mar-24"
DestSht.Cells(19, 2).Value = Mar24
DestSht.Cells(19, 3).Value = strtMar24 + progMar24 + compMar24
DestSht.Cells(19, 4).Value = progMar24 + compMar24
DestSht.Cells(19, 5).Value = compMar24
'DestSht.Cells(7, 15).Value = pubMar23

DestSht.Cells(20, 1).Value = "1-Apr-24"
DestSht.Cells(20, 2).Value = Apr24
DestSht.Cells(20, 3).Value = strtApr24 + progApr24 + compApr24
DestSht.Cells(20, 4).Value = progApr24 + compApr24
DestSht.Cells(20, 5).Value = compApr24
'DestSht.Cells(8, 15).Value = pubApr24

DestSht.Cells(21, 1).Value = "1-May-24"
DestSht.Cells(21, 2).Value = May24
DestSht.Cells(21, 3).Value = strtMay24 + progMay24 + compMay24
DestSht.Cells(21, 4).Value = progMay24 + compMay24
DestSht.Cells(21, 5).Value = compMay24
'DestSht.Cells(9, 15).Value = pubMay24

DestSht.Cells(22, 1).Value = "1-Jun-24"
DestSht.Cells(22, 2).Value = Jun24
DestSht.Cells(22, 3).Value = strtJun24 + progJun24 + compJun24
DestSht.Cells(22, 4).Value = progJun24 + compJun24
DestSht.Cells(22, 5).Value = compJun24
'DestSht.Cells(10, 15).Value = pubJun24

DestSht.Cells(23, 1).Value = "1-Jul-24"
DestSht.Cells(23, 2).Value = Jul24
DestSht.Cells(23, 3).Value = strtJul24 + progJul24 + compJul24
DestSht.Cells(23, 4).Value = progJul24 + compJul24
DestSht.Cells(23, 5).Value = compJul24
'DestSht.Cells(11, 15).Value = pubJul24

DestSht.Cells(24, 1).Value = "1-Aug-24"
DestSht.Cells(24, 2).Value = Aug24
DestSht.Cells(24, 3).Value = strtAug24 + progAug24 + compAug24
DestSht.Cells(24, 4).Value = progAug24 + compAug24
DestSht.Cells(24, 5).Value = compAug24
'DestSht.Cells(12, 15).Value = pubAug24

DestSht.Cells(25, 1).Value = "1-Sep-24"
DestSht.Cells(25, 2).Value = Sep24
DestSht.Cells(25, 3).Value = strtSep24 + progSep24 + compSep24
DestSht.Cells(25, 4).Value = progSep24 + compSep24
DestSht.Cells(25, 5).Value = compSep24
'DestSht.Cells(13, 15).Value = pubSep24

DestSht.Cells(26, 1).Value = "1-Oct-24"
DestSht.Cells(26, 2).Value = Oct24
DestSht.Cells(26, 3).Value = strtOct24 + progOct24 + compOct24
DestSht.Cells(26, 4).Value = progOct24 + compOct24
DestSht.Cells(26, 5).Value = compOct24
'DestSht.Cells(14, 15).Value = pubOct24

DestSht.Cells(27, 1).Value = "1-Nov-24"
DestSht.Cells(27, 2).Value = Nov24
DestSht.Cells(27, 3).Value = strtNov24 + progNov24 + compNov24
DestSht.Cells(27, 4).Value = progNov24 + compNov24
DestSht.Cells(27, 5).Value = compNov24
'DestSht.Cells(15, 15).Value = pubNov24

DestSht.Cells(28, 1).Value = "1-Dec-24"
DestSht.Cells(28, 2).Value = Dec24
DestSht.Cells(28, 3).Value = strtDec24 + progDec24 + compDec24
DestSht.Cells(28, 4).Value = progDec24 + compDec24
DestSht.Cells(28, 5).Value = compDec24

DestSht.Cells(29, 1).Value = "1-Jan-25"
DestSht.Cells(29, 2).Value = Jan25
DestSht.Cells(29, 3).Value = strtJan25 + progJan25 + compJan25
DestSht.Cells(29, 4).Value = progJan25 + compJan25
DestSht.Cells(29, 5).Value = compJan25
'DestSht.Cells(5, 15).Value = pubJan23

DestSht.Cells(30, 1).Value = "1-Feb-25"
DestSht.Cells(30, 2).Value = Feb25
DestSht.Cells(30, 3).Value = strtFeb25 + progFeb25 + compFeb25
DestSht.Cells(30, 4).Value = progFeb25 + compFeb25
DestSht.Cells(30, 5).Value = compFeb25
'DestSht.Cells(6, 15).Value = pubFeb23

DestSht.Cells(31, 1).Value = "1-Mar-25"
DestSht.Cells(31, 2).Value = Mar25
DestSht.Cells(31, 3).Value = strtMar25 + progMar25 + compMar25
DestSht.Cells(31, 4).Value = progMar25 + compMar25
DestSht.Cells(31, 5).Value = compMar25
'DestSht.Cells(7, 15).Value = pubMar23

DestSht.Cells(32, 1).Value = "1-Apr-25"
DestSht.Cells(32, 2).Value = Apr25
DestSht.Cells(32, 3).Value = strtApr25 + progApr25 + compApr25
DestSht.Cells(32, 4).Value = progApr25 + compApr25
DestSht.Cells(32, 5).Value = compApr25
'DestSht.Cells(8, 15).Value = pubApr25

DestSht.Cells(33, 1).Value = "1-May-25"
DestSht.Cells(33, 2).Value = May25
DestSht.Cells(33, 3).Value = strtMay25 + progMay25 + compMay25
DestSht.Cells(33, 4).Value = progMay25 + compMay25
DestSht.Cells(33, 5).Value = compMay25
'DestSht.Cells(9, 15).Value = pubMay25

DestSht.Cells(34, 1).Value = "1-Jun-25"
DestSht.Cells(34, 2).Value = Jun25
DestSht.Cells(34, 3).Value = strtJun25 + progJun25 + compJun25
DestSht.Cells(34, 4).Value = progJun25 + compJun25
DestSht.Cells(34, 5).Value = compJun25
'DestSht.Cells(10, 15).Value = pubJun25

DestSht.Cells(35, 1).Value = "1-Jul-25"
DestSht.Cells(35, 2).Value = Jul25
DestSht.Cells(35, 3).Value = strtJul25 + progJul25 + compJul25
DestSht.Cells(35, 4).Value = progJul25 + compJul25
DestSht.Cells(35, 5).Value = compJul25
'DestSht.Cells(11, 15).Value = pubJul25

DestSht.Cells(36, 1).Value = "1-Aug-25"
DestSht.Cells(36, 2).Value = Aug25
DestSht.Cells(36, 3).Value = strtAug25 + progAug25 + compAug25
DestSht.Cells(36, 4).Value = progAug25 + compAug25
DestSht.Cells(36, 5).Value = compAug25
'DestSht.Cells(12, 15).Value = pubAug25

DestSht.Cells(37, 1).Value = "1-Sep-25"
DestSht.Cells(37, 2).Value = Sep25
DestSht.Cells(37, 3).Value = strtSep25 + progSep25 + compSep25
DestSht.Cells(37, 4).Value = progSep25 + compSep25
DestSht.Cells(37, 5).Value = compSep25
'DestSht.Cells(13, 15).Value = pubSep25

DestSht.Cells(38, 1).Value = "1-Oct-25"
DestSht.Cells(38, 2).Value = Oct25
DestSht.Cells(38, 3).Value = strtOct25 + progOct25 + compOct25
DestSht.Cells(38, 4).Value = progOct25 + compOct25
DestSht.Cells(38, 5).Value = compOct25
'DestSht.Cells(14, 15).Value = pubOct25

DestSht.Cells(39, 1).Value = "1-Nov-25"
DestSht.Cells(39, 2).Value = Nov25
DestSht.Cells(39, 3).Value = strtNov25 + progNov25 + compNov25
DestSht.Cells(39, 4).Value = progNov25 + compNov25
DestSht.Cells(39, 5).Value = compNov25
'DestSht.Cells(15, 15).Value = pubNov25

DestSht.Cells(40, 1).Value = "1-Dec-25"
DestSht.Cells(40, 2).Value = Dec25
DestSht.Cells(40, 3).Value = strtDec25 + compDec25 + progDec25
DestSht.Cells(40, 4).Value = progDec25 + compDec25
DestSht.Cells(40, 5).Value = compDec25

Sheets("Cumulative Test Chart").Activate
'Sheets("Test OverView").Visible = False


DestSht.Cells(42, 11).Value = Untest


End Sub


