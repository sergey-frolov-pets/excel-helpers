Attribute VB_Name = "MClipboard"
'@Folder("sfSnippets")

''' <summary>
''' --------------------------
''' Module <c>MClipboard.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' procedures to work with Clipboard
''' --------------------------
''' <source>
''' https://bytecomb.com/copy-and-paste-in-vba/
''' </source>
''' provided by Sergey Frolov (pet-projects@sergey-frolov.ru)
''' --------------------------
''' <list>
'''     Sub <c>SetTextToClipboard</c> - Copy text to Clipboard
'''     Function <c>GetTextFromClipboard</c> - Paste text from Clipboard to variable
''' </list>
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

''' <const><c>DATAOBJECT_BINDING</c> - Clipboard object ID</const>
Const DATAOBJECT_BINDING As String = "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"

''' <summary>
''' --------------------------
''' Sub <c>setTextToClipboard</c>
''' --------------------------
''' Copy text to Clipboard
''' --------------------------
''' <param><c>text</c> - text to put into Clipboard</param>
''' --------------------------
''' </summary>
Public Sub setTextToClipboard(ByVal text As String)
    
    With CreateObject(DATAOBJECT_BINDING)
        .SetText text
        .PutInClipboard
    
    End With

End Sub
 
''' <summary>
''' --------------------------
''' Function <c>getTextFromClipboard</c>
''' --------------------------
''' Get text from Clipboard
''' --------------------------
''' <returns>Text from Clipboard</returns>
''' </summary>
Public Function getTextFromClipboard() As String
    
    With CreateObject(DATAOBJECT_BINDING)
        .GetFromClipboard
        getTextFromClipboard = .GetText
    
    End With

End Function
