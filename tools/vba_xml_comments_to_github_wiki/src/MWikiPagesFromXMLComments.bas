Attribute VB_Name = "MWikiPagesFromXMLComments"
'@Folder("sfXMLComments")
''' <summary>
''' --------------------------
''' Module <c>MWikiPagesFromXMLComments.bas</c>
''' --------------------------
''' Creates Github wiki pages from XML Comments in VBA files(*.bas, *.cls, *.frm)
''' This tool is based on Microsoft recommendations for documenting VBA code:
''' https://docs.microsoft.com/en-us/dotnet/visual-basic/reference/language-specification/documentation-comments
''' --------------------------
''' <references>
''' <c>MSugar.bas</c>
''' <c>MText.bas</c>
''' <c>CWikiCommentFunction.cls</c>
''' <c>CWikiCommentModule.cls</c>
''' <c>CWikiCommentSub.cls</c>
''' <c>CWikiModulePage.cls</c>
''' <c>IWikiComment.cls</c>
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

''' <summary>
''' --------------------------
''' Sub <c>ChooseVBAFolder_Click</c>
''' --------------------------
''' version 0.1 (2020-09-03)
''' --------------------------
''' Handler for button to select VBA folder
''' --------------------------
''' <param><c>None</c></param>
''' --------------------------
''' </summary>
Public Sub ChooseVBAFolder_Click()
    [F2] = chooseFileOrFolder("Select folder with VBA code files", True, , "*.*")
End Sub

''' <summary>
''' --------------------------
''' Sub <c>ChooseWikiFolder_Click</c>
''' --------------------------
''' version 0.1 (2020-09-03)
''' --------------------------
''' Handler for button to select Wiki-files folder
''' --------------------------
''' <param><c>None</c></param>
''' --------------------------
''' </summary>
Public Sub ChooseWikiFolder_Click()
    [F5] = chooseFileOrFolder("Select folder were wiki-files (*.md) will be saved", True, , "*.*")
End Sub

''' <summary>
''' --------------------------
''' Sub <c>ExtractComments_Click</c>
''' --------------------------
''' version 0.1 (2020-09-03)
''' --------------------------
''' Handler for button to start XML-comments extraction
''' --------------------------
''' <param><c>None</c></param>
''' --------------------------
''' </summary>
Public Sub ExtractComments_Click()
   Dim files As Integer
   
   files = CreateWikiPagesFromSelectedFolder([F2], [F5])
   
   If files > 0 Then
        MsgBox files & " wiki-file(s) were created successfuly.", vbInformation, "Done"
        
        ActiveSheet.Hyperlinks.Add Anchor:=[F8], Address:=[F5], TextToDisplay:="Extracted to " & [F5]
    Else
        [F8] = "VBA files with XML comments were not found in [" & [F2] & "]"
        MsgBox [F8], vbExclamation, "Check the source folder or files"
        If ActiveSheet.Hyperlinks.count > 0 Then ActiveSheet.Hyperlinks(1).Delete
    End If

End Sub

''' <summary>
''' --------------------------
''' Function <c>CreateWikiPagesFromSelectedFolder</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Find all VBA files in source folder (*.bas, *.cls, *.frm)
''' Extract XML comments from it
''' Find all VBA files in source folder (*.bas, *.cls, *.frm)
''' --------------------------
'''<returns>Number of extracted and successfuly converted VBA-files</returns>
''' --------------------------
''' <param><c>sourceFolder</c> - Source folder with files exported from VBA project which contain XML Comments</param>
''' <param><c>targetFolder</c> - Target folder where *.md wiki-files will be stored</param>
''' --------------------------
''' </summary>
Public Function CreateWikiPagesFromSelectedFolder(ByVal sourceFolder As String, Optional targetFolder As String) As Integer
    Dim files As Variant
    Dim file As Variant
    
    If sourceFolder = vbNullString Then sourceFolder = ActiveWorkbook.Path
    If targetFolder = vbNullString Then targetFolder = sourceFolder

    files = getFilesList(sourceFolder, "*.bas;*.cls;*.frm", False)
    
    For Each file In files
        If file <> vbNullString Then createWikiPageFromVBAFileWithXMLComments file, targetFolder & "\" & Dir(file) & ".md"
    Next
    
    CreateWikiPagesFromSelectedFolder = UBound(files)
End Function

''' <summary>
''' --------------------------
''' Sub <c>createWikiPageFromVBAFileWithXMLComments</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Extract XML comments from source file and
''' and convert them to Github wiki-file as target file name
''' By default, wiki-file name will be in format [source_vba_file.ext.md] (i.e. Main.bas -> Main.bas.md)
''' --------------------------
''' <param><c>fromVBAFileFullPath</c> - Full path to source file with XML comments</param>
''' <param><c>toWikiFileName</c> - Optional. Full path to target file in Github wiki format. By default, wiki-file name will be in format [source_vba_file.ext.md] (i.e. Main.bas -> Main.bas.md)</param>
''' --------------------------
''' </summary>
Public Sub createWikiPageFromVBAFileWithXMLComments(ByVal fromVBAFileFullPath As String, Optional toWikiFileName As String)
    Dim fileContent  As String
    Dim fileLines    As Variant
    Dim comments     As Variant
    Dim wikicomments As Variant
    Dim fileLine     As Variant
    Dim curComment   As Variant
    Dim wikiPage     As CWikiModulePage
    Dim i            As Integer
    
    fileContent = loadFileToString(fromVBAFileFullPath)
    fileLines = Split(fileContent, Chr(10), , vbTextCompare)

    Dim isXMLComment As Boolean

    curComment = vbNullString
    ReDim comments(0)

    For Each fileLine In fileLines
        If Mid(fileLine, 1, 3) = "'''" Then
            isXMLComment = True
            If Len(fileLine) > 3 Then
                fileLine = Mid(fileLine, 4)
                If Mid(fileLine, 1, 1) = " " Then fileLine = Mid(fileLine, 2)
                incr curComment, fileLine & Chr(10)
            Else
                incr curComment, Chr(10)
            End If
        Else
            isXMLComment = False
            If curComment <> vbNullString Then addToArray comments, curComment
            curComment = vbNullString
        End If
    Next

    If curComment <> vbNullString Then addToArray comments, curComment

    Set wikiPage = New CWikiModulePage

    For i = 1 To UBound(comments)
        wikiPage.addCommentFromString comments(i)
    Next

    If toWikiFileName = vbNullString Then
        saveStringToFile wikiPage.saveToString(), fromVBAFileFullPath & ".md"
    Else
        saveStringToFile wikiPage.saveToString(), toWikiFileName

    End If

End Sub

