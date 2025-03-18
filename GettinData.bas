Attribute VB_Name = "GettinData"
Sub Choose_Source()
' This will be the beginning to pick a folder

Dim fldr As FileDialog
Dim sItem As String

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
.Title = "Select a Folder"
.AllowMultiSelect = False
.InitialFileName = strPath

If .Show <> -1 Then GoTo NextCode

sItem = .SelectedItems(1)

End With


NextCode:
GetFolder = sItem
Set fldr = Nothing


'Create array
    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sItem)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then GoTo SkipEnd

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
    
    xfile = InStr(oFile.Name, ".xlsx")      'look for excel file?
    
    If xfile > 0 Then
        vaArray(i) = oFile.Name
        i = i + 1
       
     End If
     
    Next
    For k = 1 To oFiles.Count
    
        Cells(k, 1).Value = vaArray(k)
        
    Next
    'listfiles = vaArray


'Dim x() As Variant

'x = listfiles(sItem)

'For i = 0 To x.Count

'Next
SkipEnd:

End Sub

Function listfiles(ByVal sPath As String)

    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Function

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next

    listfiles = vaArray

End Function
