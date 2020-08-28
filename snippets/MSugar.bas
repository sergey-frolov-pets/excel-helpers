Attribute VB_Name = "MSugar"
'@Folder("sfSnippets")

''' <summary>
''' --------------------------
''' Module <c>MSugar.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' support functions and procedures to simplify VBA coding
''' --------------------------
''' <list>
'''     <c>Incr</c> - Replaces code like i=i+1, or i=i+value
'''     <c>AddToArray</c> - Add new element to the end of dynamic array
'''     <c>RemoveFromArray</c> - Remove element with target index from dynamic array
''' </list>
''' --------------------------
''' <references>
'''<c>Module/Class name</c>
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

''' <summary>
''' --------------------------
''' Sub <c>addToArray</c>
''' --------------------------
''' Add new element to the end of dynamic array
''' array starts from index 1
''' --------------------------
''' <example>
''' For example:
''' <code>
'''     Redim a(0)
'''     addToArray a, "test"
''' </code>
''' Results:
'''     a(1) = "test"
''' </example>
''' --------------------------
''' <param><c>arr</c> - target array</param>
''' <param><c>newElement</c> - new element</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub addToArray(arr As Variant, ByVal newElement As Variant)
    
    ReDim Preserve arr(UBound(arr) + 1)
    
    If IsObject(newElement) Then
        Set arr(UBound(arr)) = newElement
    
    Else
        arr(UBound(arr)) = newElement
    
    End If
    
End Sub

''' <summary>
''' --------------------------
''' Sub <c>removeFromArray</c>
''' --------------------------
''' Remove element with target index from dynamic array
''' --------------------------
''' <param><c>arr</c> - target array</param>
''' <param><c>Index</c> - target index of the element to be deleted</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Sub RemoveFromArray(arr As Variant, ByVal Index As Integer)
    Dim i As Integer
    
    If i < UBound(arr) Then
        For i = Index To UBound(arr) - 1
            If IsObject(arr(i)) Then
                Set arr(i) = arr(i + 1)
            
            Else
                arr(i) = arr(i + 1)
            
            End If
        Next
    
    End If
    
    ReDim Preserve arr(UBound(arr) - 1)

End Sub

