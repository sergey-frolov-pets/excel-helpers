Attribute VB_Name = "MSupport"
'@Folder("sfProfiler")
'------------------------------------------------------------------------------
'Support functions
'------------------------------------------------------------------------------

Public Sub incr(toVariable As Variant, Optional addValue As Variant = 1)
    toVariable = toVariable + addValue

End Sub

Public Function fileExists(ByVal fullPathFileName As String) As Boolean
    
    fileExists = IIf(Len(Dir(fullPathFileName)) = 0, False, True)

End Function

Public Function loadFileToString(ByVal fullPathFileName As String) As String
Dim iFile As Integer

    On Error GoTo Bad
    
    iFile = FreeFile
    Open fullPathFileName For Input As #iFile
    loadFileToString = Input(LOF(iFile), iFile)
    Close #iFile

Exit Function

Bad:
    loadFileToString = ""

End Function


Public Sub addToArray(arr As Variant, ByVal newElement As Variant)
    
    ReDim Preserve arr(UBound(arr) + 1)
    
    If IsObject(newElement) Then
        Set arr(UBound(arr)) = newElement
    
    Else
        arr(UBound(arr)) = newElement
    
    End If
    
End Sub

Public Sub RemoveFromArray(arr As Variant, ByVal index As Integer)
    Dim i As Integer
    If i < UBound(arr) Then
        For i = index To UBound(arr) - 1
            If IsObject(arr(i + 1)) Then
                Set arr(i) = arr(i + 1)
            Else
                arr(i) = arr(i + 1)
            
            End If
        Next
    
    End If
    
    If UBound(arr) > 0 Then
        ReDim Preserve arr(UBound(arr) - 1)
    Else
        ReDim arr(0)
    End If
End Sub

Public Function fromTemplate(template As String, ParamArray values() As Variant) As String
Dim cur As String, i As Integer
cur = template
    
For i = 0 To UBound(values)
    cur = Replace(cur, "{%" & i + 1 & "%}", values(i))
Next
fromTemplate = cur
End Function

Public Function columnNameByIndex(ByVal columnIndex As Long) As String
    Dim vArr
    
    vArr = Split(Cells(1, columnIndex).Address(True, False), "$")
    
    columnNameByIndex = vArr(0)
End Function

Public Function lastColumnInTheRow(ByVal forRow As Integer, Optional inSheet As Variant, Optional nonEmptyCellValue As Boolean = False) As Integer
    Dim sht, i As Integer
    
    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            Set sht = Sheets(inSheet)
        End If
    
    End If

    lastColumnInTheRow = sht.Cells(forRow, sht.columns.count).End(xlToLeft).Column

    If nonEmptyCellValue Then
        For i = lastColumnInTheRow To 1 Step -1
            lastColumnInTheRow = i
            If sht.Cells(forRow, i).value <> "" Then Exit For
        Next
    
    End If
    Set sht = Nothing
End Function

Public Function lastRowInColumn(inColumn As String, Optional inSheet As Variant) As Integer
    Dim sht
    
    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    ElseIf IsObject(inSheet) Then
        Set sht = inSheet
    
    Else
        Set sht = Sheets(inSheet)
    
    End If

    lastRowInColumn = sht.Range(inColumn & "1").Cells(65536, inColumn).End(xlUp).row

    Set sht = Nothing
End Function

Public Sub enableFastCode(set_fast As Boolean)
    On Error Resume Next
    
    With Application
        If set_fast Then
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .DisplayStatusBar = False
            .StatusBar = False
            .DisplayAlerts = False
            
        Else
            .EnableEvents = True
            .DisplayStatusBar = True
            .StatusBar = True
            .DisplayAlerts = True
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
        End If
    
    End With

End Sub

Public Function countSubStrings(ByVal inString As String, ByVal subString As String) As Long
countSubStrings = (Len(inString) - Len(Replace(inString, subString, ""))) / Len(subString)
End Function

