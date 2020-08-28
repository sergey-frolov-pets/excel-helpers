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

''' <summary>
''' --------------------------
''' Sub <c>log</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Copy/paste the Row of selected start cell
''' and put there time, reporter, action and comments starting from start Cell
'''
''' <example>
''' For example:
''' <code>
'''     log "first email recieved", "admin", "letter from the Customer", [A1]
'''
'''     initLogging [A1]
'''     log "second email recieved", "admin"
''' </code>
''' Results:
'''   28.08.2020 13:15|second email recieved|admin
'''   28.08.2020 13:00|first email recieved |admin|letter from the Customer
''' </example>
''' --------------------------
''' <param><c>action</c> - Action to be logged</param>
''' <param><c>who</c> - Reporter name or ID</param>
''' <param><c>comments</c> - Comments for the action</param>
''' <param><c>startCell</c> - Target cell</param>
''' --------------------------
''' </summary>

Public Sub log(ByVal action As String, Optional who As String, Optional comments As String, Optional startCell As Range)
    Dim topRow As Integer
    
    enableFastCode True
        
        If IsMissing(startCell) Then
            If IsMissing(iStartFromCell) Then
                Err.Raise 501, "MLogging.log()", "Starting cell not stated. Run initLogging() first."
            Else
                Set startCell = iStartFromCell
            End If
        End If
        
        topRow = startCell.Row
        
        cloneRow topRow, , , , True
        putValuesToRow iStartFromCell, Format(Now(), "dd.mm.yy hh:mm"), who, action, comments
        Application.CutCopyMode = False
        
        '[A1].Select 'TODO Change this cell address to more suitable for you
    
    enableFastCode False
End Sub



