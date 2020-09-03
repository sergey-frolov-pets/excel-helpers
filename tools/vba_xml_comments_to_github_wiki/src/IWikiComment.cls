VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWikiComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfXMLComments")
''' <summary>
''' --------------------------
''' Module <c>IWikiComment.cls</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Interface for different types of XML comments (Sub and  Function)
''' --------------------------
''' <references>
''' <c>none</c>
''' </references>
''' --------------------------
''' created 2020-09-02
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

Public Function name() As String

End Function

Public Function description() As String

End Function

Public Function version() As String

End Function

Public Sub loadFromString(ByVal commentStrings As String)

End Sub

Public Function saveToString() As String

End Function
