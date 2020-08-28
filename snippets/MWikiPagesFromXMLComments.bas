Attribute VB_Name = "MWikiPagesFromXMLComments"
'@Folder("sfSnippets")
''' <summary>
''' --------------------------
''' Module <c>MWikiPagesFromXMLComments</c>
''' --------------------------
''' Creates Github wiki pages from XML Comments in VBA files(*.bas, *.cls, *.frm)
''' --------------------------
''' <references>
''' <c>MSugar.bas</c>
''' <c>MText.bas</c>
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

Public Sub test()
createWikiPageFromVBAFileWithXMLComments "C:\GitHub\excel-helpers\snippets\MSugar.bas"

End Sub

Public Sub createWikiPageFromVBAFileWithXMLComments(ByVal fromVBAFileFullPath As String, Optional toWikiFileName As String = vbNullString)
Dim fileContent As String
Dim fileLines As Variant
Dim comments As Variant
Dim fileLine  As Variant
Dim curComment As Variant

fileContent = loadFileToString(fromVBAFileFullPath)
fileLines = Split(fileContent, Chr(10), , vbTextCompare)

Dim isXMLComment As Boolean

curComment = vbNullString
ReDim comments(0)

For Each fileLine In fileLines
    If Mid(fileLine, 1, 3) = "'''" Then
        isXMLComment = True
        If Len(fileLine) > 3 Then
            incr curComment, Trim(Mid(fileLine, 4)) & Chr(10)
        End If
    Else
        isXMLComment = False
        If curComment <> vbNullString Then addToArray comments, curComment
        curComment = vbNullString
    End If
Next

If curComment <> vbNullString Then addToArray comments, curComment

fileLine = vbNullString
For Each curComment In comments
    incr fileLine, curComment & Chr(10)
Next

saveStringToFile fileLine, "C:\GitHub\excel-helpers\snippets\MSugar.bas.txt"

End Sub

Public Sub createFolderWithWikiPagesFromFolderWithVBAFiles(ByVal fromVBAFileFullPath As String, Optional toWikiFileName As String = vbNullString)

End Sub

Public Function loadFileToString(ByVal fileNameFullPath As String) As String
Dim iFile As Integer
iFile = FreeFile

Open fileNameFullPath For Input As #iFile
loadFileToString = Input(LOF(iFile), iFile)
Close #iFile

End Function

Public Sub saveStringToFile(ByVal strContent As String, ByVal fileNameFullPath As String)
    Dim iFile As Integer
    iFile = FreeFile
    
    Open fileNameFullPath For Output As #iFile
    Print #iFile, strContent
    Close #iFile
End Sub
