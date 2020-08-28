Attribute VB_Name = "MTools"
'@Folder("sfSnippets")

''' <summary>
''' --------------------------
''' Module <c>MTools.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' contains simple tools for daily business-tasks
''' --------------------------
''' <list>
'''     <c>FindDublicates</c> - Find rows in selected Range with similar values in selected Columns
''' </list>
''' --------------------------
''' <references>
''' <c>MSugar.bas</c>
''' <c>MClipboard.bas</c>
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
''' Sub <c>FindDublicates</c>
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
Public Sub FindDublicates(inRange As Range, idInColumns As String)
Dim i As Long, j As Long, k As Integer
Dim columnsID As Variant
Dim res As String, finRes As String, curID As String, curChk As String
columnsID = Split(idInColumns, ",")
inDubles = " "

For i = 1 To inRange.Rows.count - 1
    curR = inRange.Row + i - 1
    curID = ""
    
    For j = 0 To UBound(columnsID)
        Incr curID, inRange.Parent.Range(columnsID(j) & inRange.Row + i - 1).Value & " "
    Next
    
    res = ""
    For j = i + 1 To inRange.Rows.count
        If InStr(inDubles, " " & j & " ") = 0 Then
            curChk = ""
            For k = 0 To UBound(columnsID)
                Incr curChk, inRange.Parent.Range(columnsID(k) & inRange.Row + j - 1).Value & " "
            Next
           If curChk = curID Then
                Incr inDubles, j & " "
                If res = "" Then Incr res, inRange.Row + i - 1 & vbTab
                Incr res, inRange.Row + j - 1 & vbTab
            End If
        End If
    Next
    If res <> "" Then Incr finRes, vbCrLf & res
Next
SetTextToClipboard finRes
If res <> "" Then MsgBox "Dublicates' row numbers were copied to Clipboard" Else MsgBox "Dublicates not found"
End Sub

