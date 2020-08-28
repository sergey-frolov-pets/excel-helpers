Attribute VB_Name = "MLogging"
'@Folder("sfSnippets")
''' <summary>
''' --------------------------
''' Module <c>MLogging.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' logging events to Excel spreadsheet using VBA
''' --------------------------
''' <references>
''' <c>MAcessors.bas</c>
''' <c>MSugar.bas</c>
''' </references>
''' --------------------------
''' created 2020-08-26
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

''' <param><c>iStartFromCell</c> - Local variable to save starting cell for logging</param>
Private iStartFromCell As Range

''' <summary>
''' --------------------------
''' Sub <c>initLogging</c>
''' --------------------------
''' Set starting cell for logging. May be extended with additional params.
''' --------------------------
''' <param><c>Name</c> - Description</param>
''' --------------------------
''' created 2020-08-26
''' --------------------------
''' </summary>
Public Sub initLogging(startCell As Range)
    Set iStartFromCell = startCell

End Sub

Public Sub log(ByVal action As String, Optional who As String, Optional comments As String)
    Dim topRow As Integer
    
    enableFastCode True
    
        If IsMissing(iStartFromCell) Then Err.Raise 501, "MLogging.log()", "Starting cell not stated. Run initLogging() first."
        
        topRow = iStartFromCell.Row
        
        cloneRow topRow, , , , True
        putValuesToRow iStartFromCell, Format(Now(), "dd.mm.yy hh:mm"), who, action, comments
        Application.CutCopyMode = False
        
        [A1].Select 'TODO Change this cell address to more suitable for you
    
    enableFastCode False
End Sub



