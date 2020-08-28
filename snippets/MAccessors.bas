Attribute VB_Name = "MAccessors"
'@Folder("sfSnippets")
''' <summary>
''' --------------------------
''' Module <c>MAccessors.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' functions and procedures to simplify access to Excel objects like workbooks, sheets, rows, columns, etc.
''' --------------------------
''' <references>
''' <c>MSugar.bas</c>
''' </references>
''' --------------------------
''' created 2020-08-20
''' by Sergey Frolov (pet-projects@sergey-frolov.ru)
''' --------------------------
''' </summary>
'''
''' <license>
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see
''' https://www.gnu.org/licenses/
''' </license>

Option Explicit

''' <summary>
''' --------------------------
''' Function <c>getCellValueWithoutOpeningFile</c>
''' --------------------------
''' Get cell value without opening external Excel-file
''' --------------------------
''' <returns>Value of cell</returns>
''' --------------------------
''' <param><c>fileName</c> - Name of external Excel file</param>
''' <param><c>sheetName</c> - target Sheet name</param>
''' <param><c>rangeAddress</c> - Range name, where the first cell is a point of our interest</param>
''' <param><c>folderOfExcelFile</c> - Optional. The name of folder where the target file is located.
''' By default current folder will be applied</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function getCellValueWithoutOpeningFile(ByVal fileName As String, ByVal sheetName As String, _
                                               ByVal rangeAddress As String, Optional folderOfExcelFile As String = "") As Variant
    Dim arg As String
    Dim fl As String

    If folderOfExcelFile = "" Then
        fl = Mid(fileName, InStrRev(fileName, "\") + 1)
        arg = "'" & Left(fileName, Len(fileName) - Len(fl)) & _
              "[" & fl & "]" & sheetName & "'!" & _
              Range(rangeAddress).Range("A1").Address(, , xlR1C1)

    Else
        If Right(folderOfExcelFile, 1) <> "\" Then incr folderOfExcelFile, "\"
    
        If Dir(folderOfExcelFile & fileName) = vbNullString Then
            getCellValueWithoutOpeningFile = "File Not Found"
            Exit Function
        End If
    
        arg = "'" & folderOfExcelFile & "[" & fileName & "]" & _
              sheetName & "'!" & Range(rangeAddress).Range("A1").Address(, , xlR1C1)
    
    End If

    getCellValueWithoutOpeningFile = ExecuteExcel4Macro(arg)
End Function

''' <summary>
''' --------------------------
''' Function <c>findRow</c>
''' --------------------------
''' Find the first row in the target range with the same values
''' as in the source row for specified columns
''' --------------------------
'''<returns>Row object</returns>
''' --------------------------
''' <param><c>rowToFind</c> - source/example Row</param>
''' <param><c>columnsIDsInRow</c> - list of the column names of the source row separated by coma which will be used for search</param>
''' <param><c>inRange</c> - target Range, where to search</param>
''' <param><c>columnsIDsInRange</c> - list of the column names of target range separated by coma which will be used for search
''' Column names can be different for the source and target ranges and will be compared by the order
''' in <c>columnsIDsInRow</c> and <c>columnsIDsInRange</c> parameters.</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function findRow(rowToFind As Range, columnsIDsInRow As String, inRange As Range, columnsIDsInRange As String) As Range

    Dim i As Long, j As Long, shift As Long
    Dim columnsID As Variant, columnsToCheckID As Variant
    Dim lastRowInRange As Integer

    columnsID = Split(columnsIDsInRow, ",")
    columnsToCheckID = Split(columnsIDsInRange, ",")

    shift = 0
    lastRowInRange = inRange.Row + inRange.Rows.count - 1

mrkNext:

    i = match(rowToFind.Parent.Range(columnsID(0) & rowToFind.Row), inRange.Parent.Range(columnsToCheckID(0) & inRange.Row + shift & ":" & columnsToCheckID(0) & lastRowInRange))

    If i = 0 Or shift > lastRowInRange Then
        Set findRow = Nothing
        Exit Function
    Else
        For j = 1 To UBound(columnsID)
            If inRange.Parent.Range(columnsToCheckID(j) & inRange.Row + shift + i - 1).Value <> rowToFind.Parent.Range(columnsID(j) & rowToFind.Row).Value Then
                shift = shift + i
                GoTo mrkNext
            End If
        Next
    End If

    Set findRow = inRange.Parent.Rows(shift + i + inRange.Row - 1)
End Function

''' <summary>
''' --------------------------
''' Sub <c>getArrayFromRow</c>
''' --------------------------
''' Put values from target row to target dynamic array.
''' Values can be taken from the first cell of target range to the last cell,
''' of the first row in the range, or from the first N columns only.
'''
''' <example>
''' For example:
''' <code>
'''     getArrayFromRow [C3], myArray, 3
''' </code>
''' Results:
'''
''' </example>
'''
''' --------------------------
''' <param><c>firstCell</c> - Range. Starting cell(s) for taking values</param>
''' <param><c>dataArray</c> - Target DYNAMIC array, where the values will be placed</param>
''' <param><c>length</c> - Optional. Amount of the values to be taken from the starting cell.
''' By default values will be taken to the last non-empty cell in the target row.</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub getArrayFromRow(firstCell As Range, dataArray As Variant, Optional length = 0)

    Dim i As Integer
    Dim l As Integer
    Dim s As Integer
    Dim r As Integer
    
    s = firstCell.Column
    r = firstCell.Row
    If length > 0 Then l = length + s - 1 Else l = lastColumnInTheRow(r, firstCell.Parent)

    ReDim dataArray(l - s + 1)
    For i = s To l
        dataArray(i - s + 1) = firstCell.Cells(1, i - s + 1).Value
    Next

End Sub

''' <summary>
''' --------------------------
''' Sub <c>putValuesToRow</c>
''' --------------------------
''' Put any values to row starting from target cell
''' --------------------------
''' <example>
''' For example:
''' <code>
'''     putValuesToRow [B7], i, j, ,"test", [A1]
''' </code>
''' Results:
'''   4 values i, j, "test" and [A1].Value will be placed in range [B7:F7]
'''   IMPORTAINT! Param#3 (for cell [E7]) is missed - value there will be untouched.
'''   Use Empty value [i.e. PutValuesToRow [B7], i, j, Empty ,"test", [A1]] to clear it.
''' </example>
''' --------------------------
''' <param><c>firstCell</c> - starting cell</param>
''' <param><c>dataArray</c> - ParamArray of values to be placed in row</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub putValuesToRow(firstCell As Range, ParamArray dataArray())
    Dim i As Integer
    
    For i = 0 To UBound(dataArray)
        If Not IsMissing(dataArray(i)) Then firstCell.Cells(1, i + 1).Value = dataArray(i)
    Next
End Sub

''' <summary>
''' --------------------------
''' Sub <c>putArrayToRow</c>
''' --------------------------
''' Put array values to row starting from index 1
''' --------------------------
''' <param><c>firstCell</c> - starting cell</param>
''' <param><c>dataArray</c> - Array of values to be placed in row</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub putArrayToRow(firstCell As Range, dataArray As Variant)
Dim i As Integer

If IsArray(dataArray) Then
    For i = 1 To UBound(dataArray)
        firstCell.Cells(1, i).Value = dataArray(i)
    Next
End If
End Sub

''' <summary>
''' --------------------------
''' Function <c>lastColumnInTheRow</c>
''' --------------------------
''' Find the column index of the last non-empty cell in the row
''' --------------------------
'''<returns>Column index of the last non-empty cell in the row</returns>
''' --------------------------
''' <param><c>forRow</c> - Row number</param>
''' <param><c>inSheet</c> - Optional. Can be Sheet object, Sheet name or Sheet index.
''' By default ActiveSheet is used</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function lastColumnInTheRow(ByVal forRow As Integer, Optional inSheet As Variant) As Integer
    Dim sht
    
    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            Set sht = Sheets(inSheet)
        End If
    
    End If

    lastColumnInTheRow = sht.Cells(forRow, sht.Columns.count).End(xlToLeft).Column
    
    Set sht = Nothing
End Function

''' <summary>
''' --------------------------
''' Function <c>lastColumnNameInTheRow</c>
''' --------------------------
''' Find the column name of the last non-empty cell in the row
''' --------------------------
'''<returns>Column name of the last non-empty cell in the row</returns>
''' --------------------------
''' <param><c>forRow</c> - Row number</param>
''' <param><c>inSheet</c> - Optional. Can be Sheet object, Sheet name or Sheet index.
''' By default ActiveSheet is used</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function lastColumnNameInTheRow(ByVal forRow As Integer, Optional inSheet As Variant) As String
    Dim sht

    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            Set sht = Sheets(inSheet)
        End If
    
    End If

    lastColumnNameInTheRow = columnNameByIndex(lastColumnInTheRow(forRow, sht))

    Set sht = Nothing
End Function

''' <summary>
''' --------------------------
''' Function <c>columnNameByIndex</c>
''' --------------------------
''' Get column name by column index
''' --------------------------
'''<returns>Column name</returns>
''' --------------------------
''' <param><c>columnIndex</c> - Index of the target column</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function columnNameByIndex(ByVal columnIndex As Long) As String
    Dim vArr
    
    vArr = Split(Cells(1, columnIndex).Address(True, False), "$")
    
    columnNameByIndex = vArr(0)
End Function

''' <summary>
''' --------------------------
''' Function <c>lastRowInColumn</c>
''' --------------------------
''' Find row number for the last non-empty cell in the target column
''' --------------------------
'''<returns>Row number for the last non-empty cell in the target column</returns>
''' --------------------------
''' <param><c>inColumn</c> - Column name</param>
''' <param><c>inSheet</c> - Optional. Can be Sheet object, Sheet name or Sheet index.
''' By default ActiveSheet is used</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function lastRowInColumn(inColumn As String, Optional inSheet As Variant) As Integer
    Dim sht
    
    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    ElseIf IsObject(inSheet) Then
        Set sht = inSheet
    
    Else
        Set sht = Sheets(inSheet)
    
    End If

    lastRowInColumn = sht.Range(inColumn & "1").Cells(65536, inColumn).End(xlUp).Row

    Set sht = Nothing
End Function

''' <summary>
''' --------------------------
''' Function <c>lastRowInSheet</c>
''' --------------------------
''' Find row number for the last non-empty cell in the entire sheet
''' --------------------------
'''<returns>The last row number for the last non-empty cell out of all columns in the sheet</returns>
''' --------------------------
''' <param><c>inSheet</c> - Optional. Can be Sheet object, Sheet name or Sheet index.
''' By default ActiveSheet is used</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function lastRowInSheet(Optional inSheet As Variant) As Integer
    Dim sht

    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            Set sht = Sheets(inSheet)
        End If
    
    End If

    lastRowInSheet = sht.UsedRange.Cells(1, 1).Row + sht.UsedRange.Rows.count - 1
    
    Set sht = Nothing
End Function

''' <summary>
''' --------------------------
''' Sub <c>enableFastCode</c>
''' --------------------------
''' Toggle Excel autoupdate mode - SIGNIFICANTLY speed-ups calculations!
''' --------------------------
''' <example>
''' For example:
''' <code>
'''     enableFastCode True
'''     funcBigSlowCalculationsWithSheetsUpdates()
'''     enableFastCode False
''' </code>
''' Results:
'''     Dramatically saves time on refreshing Excel workbook after each cells changes
''' </example>
''' --------------------------
''' <param><c>set_fast</c> - True for switch-off Excel autoupdate, False - for switch-on</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
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

''' <summary>
''' --------------------------
''' Sub <c>deleteVisibleRows</c>
''' --------------------------
''' Delete all non-hidden rows in the ActiveSheet
''' Can be used in combination with Filetrs
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub deleteVisibleRows()

    enableFastCode True

    Dim i As Integer
    
    For i = Columns(4).Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row To 3 Step -1
        If Rows(i).Hidden = False Then Rows(i).Delete
    Next

    enableFastCode False

End Sub

''' <summary>
''' --------------------------
''' Sub <c>cloneRow</c>
''' --------------------------
''' Copy source row and paste it before target row
''' --------------------------
''' <param><c>copyRow</c> - Source row (values will be copied from this row)</param>
''' <param><c>pasteBeforeRow</c> - Target row (values will be pasted to the row BEFORE this row)</param>
''' <param><c>inSheet</c> - Optional, ActiveSheet as default value</param>
''' <param><c>toSheet</c> - Optional, inSheet as default value</param>
''' <param><c>insertRow</c> - If True will add new row before paste, if False - will replce the row before target row with copied values</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub cloneRow(copyRow As Integer, Optional pasteBeforeRow As Integer, Optional inSheet As Variant, Optional toSheet As Variant, Optional insertRow As Boolean = False)
    Dim sht, sht2

    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            If inSheet = vbNullString Then Set sht = ActiveSheet Else Set sht = Sheets(inSheet)
        End If
    
    End If

    If IsMissing(toSheet) Then
        Set sht2 = sht
    
    Else
        If IsObject(toSheet) Then
            Set sht2 = toSheet
        Else
            If toSheet = vbNullString Then Set sht2 = sht Else Set sht2 = Sheets(toSheet)
        End If
    
    End If

    If pasteBeforeRow = 0 Then pasteBeforeRow = copyRow + 1

    sht.Activate
    sht.Rows(copyRow & ":" & copyRow).Select
    Selection.Copy

    sht2.Activate
    Rows(pasteBeforeRow & ":" & pasteBeforeRow).Select

    If insertRow Then
        Selection.Insert shift:=xlDown
    
    Else
        sht2.Paste
    
    End If
End Sub

''' <summary>
''' --------------------------
''' Function <c>match</c>
''' --------------------------
''' Wrapper for Application.WorksheetFunction.Match() function
''' --------------------------
''' <returns>Returns the relative position of an item in the Range that matches a specified value.</returns>
''' --------------------------
''' <param><c>toFind</c> - Value to find</param>
''' <param><c>inRange</c> - Range where to find</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function match(ByVal toFind As Variant, inRange As Range) As Integer
On Error GoTo Bad

match = Application.WorksheetFunction.match(toFind, inRange, 0)
Exit Function

Bad:
    match = 0
End Function

''' <summary>
''' --------------------------
''' Function <c>isOpenWorkbook</c>
''' --------------------------
''' Check if workbook open or not
''' --------------------------
'''<returns>True if open, or False if not</returns>
''' --------------------------
''' <param><c>wbName</c> - Workbook name (without path to file)</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function isOpenWorkbook(ByVal wbName As String) As Boolean
    Dim wb As Workbook

    For Each wb In Workbooks
        
        If wb.Name = wbName Then
            isOpenWorkbook = True
            Exit Function
        
        End If
    
    Next
    
    isOpenWorkbook = False
End Function

''' <summary>
''' --------------------------
''' Sub <c>deleteAllRowsBelowCell</c>
''' --------------------------
''' Check if workbook open or not
''' --------------------------
''' <returns>True if open, or False if not</returns>
''' --------------------------
''' <param><c>wbName</c> - Workbook name (without path to file)</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub deleteAllRowsBelowCell(fromRange As Range)
    Range(fromRange.Cells(1, 1), fromRange.Parent.Cells.SpecialCells(xlLastCell)).Rows = ""
End Sub

''' <summary>
''' --------------------------
''' Sub <c>findDublicates</c>
''' --------------------------
''' Find rows in selected Range with similar values in selected Columns.
''' Row numbers for rows with similar values in mentioned columns will be
''' copied to Clipboard
''' --------------------------
''' <param><c>inRange</c> - Source range</param>
''' <param><c>idInColumns</c> - Column names separated by coma
''' which together can be treated as unique ID for the row</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub findDublicates(inRange As Range, idInColumns As String)
    
    Dim i As Long, j As Long, k As Integer
    Dim columnsID As Variant
    Dim res As String, finRes As String, curID As String, curChk As String
    
    columnsID = Split(idInColumns, ",")
    inDubles = " "

    For i = 1 To inRange.Rows.count - 1
        curR = inRange.Row + i - 1
        curID = ""
    
        For j = 0 To UBound(columnsID)
            incr curID, inRange.Parent.Range(columnsID(j) & inRange.Row + i - 1).Value & " "
        Next
    
        res = ""
        For j = i + 1 To inRange.Rows.count
            If InStr(inDubles, " " & j & " ") = 0 Then
                curChk = ""
                
                For k = 0 To UBound(columnsID)
                    incr curChk, inRange.Parent.Range(columnsID(k) & inRange.Row + j - 1).Value & " "
                Next
                
                If curChk = curID Then
                    incr inDubles, j & " "
                    If res = "" Then incr res, inRange.Row + i - 1 & vbTab
                    incr res, inRange.Row + j - 1 & vbTab
                End If
            
            End If
        
        Next
        If res <> "" Then incr finRes, vbCrLf & res
    
    Next
    
    setTextToClipboard finRes
    
    If res <> "" Then MsgBox "Dublicates' row numbers were copied to Clipboard" Else MsgBox "Dublicates not found"

End Sub

''' <summary>
''' --------------------------
''' Function <c>getPreviousNonEmptyValue</c>
''' --------------------------
''' Get previous non-empty value in the column above the target cell (including the cell value), will be empty if there is no non-empty values above the target cell
''' --------------------------
''' <returns>First non-empty value from the target cell or above</returns>
''' --------------------------
''' <param><c>forCell</c> - Target cell to start with</param>
''' <param><c>inSheet</c> - Optional. Can be Sheet object, Sheet name or Sheet index.
''' By default ActiveSheet is used</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Private Function getPreviousNonEmptyValue(ByVal forCell As Range, Optional inSheet As Variant) As Variant
    Dim sht, curRow As Integer, colName As String
    Dim curValue

    If IsMissing(inSheet) Then
        Set sht = ActiveSheet
    Else
        If IsObject(inSheet) Then
            Set sht = inSheet
        Else
            Set sht = Sheets(inSheet)
        End If
    End If

    colName = columnNameByIndex(forCell.Column)
    curRow = forCell.Row

    Do While IsEmpty(curValue)
        curValue = sht.Range(colName & curRow).Value
        If curRow = 1 Then Exit Do
        incr curRow, -1
    Loop

    getPreviousNonEmptyValue = curValue
End Function


