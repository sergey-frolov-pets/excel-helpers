Attribute VB_Name = "MHelper"
Option Explicit

Public Sub addToArray(arr As Variant, ByVal newElement As Variant)
ReDim Preserve arr(UBound(arr) + 1)
If IsObject(newElement) Then
    Set arr(UBound(arr)) = newElement
Else
    arr(UBound(arr)) = newElement
End If
End Sub

Public Sub incr(toVariable As Variant, Optional addValue As Variant = 1)
toVariable = toVariable + addValue
End Sub

Public Sub enableFastCode(set_fast As Boolean)
    On Error Resume Next
    With Application
       If set_fast Then
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayStatusBar = False
            .DisplayAlerts = False
        Else
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayStatusBar = True
            .DisplayAlerts = True
        End If
    End With
End Sub

Function columnNameByIndex(ByVal columnIndex As Long) As String
Dim vArr
vArr = Split(Cells(1, columnIndex).Address(True, False), "$")
columnNameByIndex = vArr(0)
End Function

Public Property Get getFilesList(ByVal folderName As String, Optional fileExt As String = "*", Optional inclSubdirectories As Boolean = False)
Dim arrFileNames() As String
Dim arrDirNames() As String
Dim varDirectory As Variant
Dim flag As Boolean
Dim i As Integer


ReDim arrFileNames(0)
ReDim arrDirNames(0)

flag = True
varDirectory = Dir(folderName & "\", vbDirectory)

While flag = True
    If varDirectory = "" Then
        flag = False
    Else
        If varDirectory <> "." And varDirectory <> ".." Then
            If (GetAttr(folderName + "\" + varDirectory) And vbDirectory) = vbDirectory Then
                addToArray arrDirNames, folderName + "\" + varDirectory + "\"
            End If
        End If
        varDirectory = Dir(, vbDirectory)
    End If
Wend
addToArray arrDirNames, folderName + "\"

For i = 1 To UBound(arrDirNames)
    flag = True
    varDirectory = Dir(arrDirNames(i))
    While flag = True
        If varDirectory = "" Then
            flag = False
        Else
            If fileExt <> "*" Then
                If LCase(Mid(varDirectory, Len(varDirectory) - Len(fileExt) + 1, Len(fileExt))) = LCase(fileExt) Then addToArray arrFileNames, arrDirNames(i) & varDirectory
            Else
                addToArray arrFileNames, arrDirNames(i) & varDirectory
            End If
            
            varDirectory = Dir
        End If
    Wend
Next

getFilesList = arrFileNames
End Property

Public Sub copyPasteRange(copyRange As Range, pastFromCell As Range)
copyRange.Parent.Activate
copyRange.Select
Selection.Copy

pastFromCell.Parent.Activate
pastFromCell.Activate
pastFromCell.Parent.Paste
End Sub

Public Function lastRowInColumn(inColumn As String, Optional inSheet As Variant) As Integer
Dim sht
If IsMissing(inSheet) Then Set sht = ActiveSheet Else Set sht = Sheets(inSheet)
With sht.Range(inColumn & "1")
    lastRowInColumn = .Cells(65536, .Column).End(xlUp).Row
End With
Set sht = Nothing
End Function

Public Sub deleteAllRowsBelowCell(fromRange As Range)
    Range(fromRange.Cells(1, 1), fromRange.Parent.Cells.SpecialCells(xlLastCell)).Rows = ""
End Sub
