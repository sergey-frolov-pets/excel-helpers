Attribute VB_Name = "MSupport"
'@Folder("OmniLogger framework")

''' <summary>
''' --------------------------
''' Module <c>MSupport.bas</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' functions and subs from https://github.com/sergey-frolov-pets/excel-helpers snippets, which are necessary for OmniLogger framework
''' --------------------------
''' created 2023-09-01
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

    lastColumnInTheRow = sht.Cells(forRow, sht.columns.Count).End(xlToLeft).Column

    If nonEmptyCellValue Then
        For i = lastColumnInTheRow To 1 Step -1
            lastColumnInTheRow = i
            If sht.Cells(forRow, i).value <> "" Then Exit For
        Next
    
    End If
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
        
    lastRowInColumn = sht.Cells(sht.rows.Count, inColumn).End(xlUp).row

    Set sht = Nothing
End Function

''' <summary>
''' --------------------------
''' Sub <c>incr</c>
''' --------------------------
''' Replaces code like i=i+1, or i=i+value
''' can be used not only for Numbers
''' <example>
''' For example:
''' <code>
'''
'''  i=1
'''  incr i ' i = 2
'''
'''  n=0.5
'''  incr n, 0.3 ' n = 0.8
'''
'''  s="Hello, "
'''  incr s, "World!" ' s = "Hello, World!"
'''
''' </code>
''' </example>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub incr(toVariable As Variant, Optional addValue As Variant = 1)
    toVariable = toVariable + addValue
End Sub
