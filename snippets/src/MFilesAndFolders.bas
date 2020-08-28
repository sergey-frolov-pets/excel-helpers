Attribute VB_Name = "MFilesAndFolders"
'@Folder("sfSnippets")

''' <summary>
''' --------------------------
''' Module <c>MFilesAndFolders.bas</c>
''' --------------------------
''' contains
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
''' Function <c>chooseFileOrFolder</c>
''' --------------------------
''' Opens dialog for selecting File or Folder
''' --------------------------
'''<returns>Full path to selected file or folder
''' or empty string if nothing was selected</returns>
''' --------------------------
''' <param><c>dialogTitle</c> - Title for dialog window</param>
''' <param><c>chooseFolder</c> - Optional. False for file selection (by default), or True for folder selection</param>
''' <param><c>initialFileName</c> - Optional. Dialog will open path to this file and select it by default</param>
''' <param><c>extensions</c> - Optional. Excel extensions by default.
''' You can use kewords: "macro" for *.xlsm files, "xlsx" for *.xlsx file, or
''' provide your extensions in format "Description1|*.ext1;Description2|*.ext2..."</param>
''' --------------------------
''' </summary>
Public Function chooseFileOrFolder(ByVal dialogTitle As String, Optional chooseFolder As Boolean = False, Optional initialFileName As String = vbNullString, Optional extensions As String = vbNullString) As String

    Dim directory As String, fileName As String, sheet As Worksheet, total As Integer, fd As Office.FileDialog
    Dim ext As Variant, i As Integer
    Dim exts As Variant
    
    
    If chooseFolder Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    Else
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    End If

    With fd
        .AllowMultiSelect = False
        If initialFileName <> vbNullString Then .initialFileName = initialFileName
        .Title = dialogTitle

        If Not chooseFolder Then
            .Filters.Clear
            
            If extensions = "macro" Then
                .Filters.Add "Excel(*.xlsm)", "*.xlsm"
                .Filters.Add "All files", "*.*"

            ElseIf extensions = "xlsx" Then
                .Filters.Add "Excel(*.xlsx)", "*.xlsx"
                .Filters.Add "All files", "*.*"

            ElseIf extensions = vbNullString Then
                .Filters.Add "Excel(*.xlsx)", "*.xlsx"
                .Filters.Add "Excel(*.xlsm)", "*.xlsm"
                .Filters.Add "Excel(*.xls)", "*.xls"
                .Filters.Add "All files", "*.*"
            Else
                ext = extensions
                exts = Split(ext, ";")
                For i = 0 To UBound(exts)
                    ext = Split(exts(i), "|")
                    If LBound(ext) = 0 And UBound(ext) = 1 Then
                        .Filters.Add ext(0), ext(1)
                    Else
                        .Filters.Add ext(0), ext(0)
                    End If
                Next
                '.filters.Add "All files", "*.*"

            End If
        
        End If
        
        If .Show = True Then fileName = .SelectedItems(1)

    End With

    chooseFileOrFolder = fileName
End Function

''' <summary>
''' --------------------------
''' Function <c>getFilesList</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Property to get the list of specified files from selected folder
''' <example>
''' For example:
''' <code>
'''     Dim myArray as variant
'''     myArray = getFilesList("C:\MyProjects\Reports\","*.docx;*.pdf",True)
''' </code>
''' Results:
'''     Array myArray() will contain the list of full paths to ALL files with extensions
'''     *.docx and *.pdf from the folder C:\MyProjects\Reports\ including subfolders.
''' </example>
''' --------------------------
'''<returns>Array of full paths for specified files</returns>
''' --------------------------
''' <param><c>folderName</c> - Target folder for file search</param>
''' <param><c>fileExtensions</c> - Optional. Function will return all files by default. List of file extensions in format "*.ext1;*.ext2;..."</param>
''' <param><c>inclSubdirectories</c> - Optional. False by default. If True - function will include sub directories to search.</param>
''' --------------------------
''' </summary>
Public Property Get getFilesList(ByVal folderName As String, Optional fileExtensions As String = "*", Optional inclSubdirectories As Boolean = False)

    Dim arrFileNames() As String
    Dim arrDirNames() As String
    Dim varDirectory As Variant
    Dim flag As Boolean
    Dim i As Integer
    Dim fileExts() As String
    Dim fileExt As Variant
    
    fileExt = Replace(fileExtensions, " ", "")
    fileExt = Replace(fileExt, "*", "")
    If fileExt = vbNullString Then fileExt = "*"
    fileExts = Split(fileExt, ";")

    ReDim arrFileNames(0)
    ReDim arrDirNames(0)

    flag = True

    If Mid(folderName, Len(folderName), 1) <> "\" Then incr folderName, "\"

    varDirectory = Dir(folderName, vbDirectory)

    While flag = True
        If varDirectory = "" Then
            flag = False
        
        Else
            If varDirectory <> "." And varDirectory <> ".." Then
                If (GetAttr(folderName + varDirectory) And vbDirectory) = vbDirectory Then
                    addToArray arrDirNames, folderName + varDirectory + "\"
                
                End If
            
            End If
            varDirectory = Dir(, vbDirectory)
        
        End If
    
    Wend
    
    addToArray arrDirNames, folderName

    If inclSubdirectories Then
        For i = 1 To UBound(arrDirNames)
            flag = True
            varDirectory = Dir(arrDirNames(i))
            
            While flag = True
                If varDirectory = "" Then
                    flag = False
                
                Else
                    
                    If fileExt <> "*" Then
                       For Each fileExt In fileExts
                            If LCase(Mid(varDirectory, Len(varDirectory) - Len(fileExt) + 1, Len(fileExt))) = LCase(fileExt) Then
                                addToArray arrFileNames, arrDirNames(i) & varDirectory
                                Exit For
                            End If
                       Next
                    Else
                       addToArray arrFileNames, arrDirNames(i) & varDirectory
                    End If
                    
                    varDirectory = Dir
                
                End If
            
            Wend
        
        Next
    
    Else
        i = UBound(arrDirNames)
        flag = True
        varDirectory = Dir(arrDirNames(i))
        
        While flag = True
            If varDirectory = "" Then
                flag = False
            
            Else
                
                If fileExt <> "*" Then
                   For Each fileExt In fileExts
                        If LCase(Mid(varDirectory, Len(varDirectory) - Len(fileExt) + 1, Len(fileExt))) = LCase(fileExt) Then
                            addToArray arrFileNames, arrDirNames(i) & varDirectory
                            Exit For
                        End If
                   Next
                Else
                   addToArray arrFileNames, arrDirNames(i) & varDirectory
                End If

                varDirectory = Dir
            
            End If
        
        Wend
    
    End If

    getFilesList = arrFileNames
End Property

''' <summary>
''' --------------------------
''' Function <c>fileExists</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Check if file exists or not
''' --------------------------
'''<returns>True if file exists, False if not.</returns>
''' --------------------------
''' <param><c>fullPathFileName</c> - Full path to the target file.</param>
''' --------------------------
''' </summary>
Public Function fileExists(ByVal fullPathFileName As String) As Boolean
    
    If Len(Dir(fullPathFileName)) = 0 Then
        fileExists = False
    
    Else
        fileExists = True
    
    End If

End Function


''' <summary>
''' --------------------------
''' Function <c>folderExists</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Check if folder exists or not
''' --------------------------
'''<returns>True if folder exists, False if not.</returns>
''' --------------------------
''' <param><c>folderFullPath</c> - Full path to the target folder.</param>
''' --------------------------
''' </summary>
Public Function folderExists(ByVal folderFullPath As String) As Boolean
    
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    folderExists = fso.folderExists(folderFullPath)

    Set fso = Nothing
End Function

''' <summary>
''' --------------------------
''' Function <c>fileCreationDate</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Get file creation date
''' --------------------------
'''<returns>Date of file creation</returns>
''' --------------------------
''' <param><c>fullPathFileName</c> - Full path to the target file</param>
''' --------------------------
''' </summary>
Public Function fileCreationDate(ByVal fullPathFileName As String) As Variant
    
    Dim fso, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(fullPathFileName)

    fileCreationDate = file.DateCreated
End Function

''' <summary>
''' --------------------------
''' Function <c>fileLastModifiedDate</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Get the last modification date for the file
''' --------------------------
'''<returns>Date of the last modification</returns>
''' --------------------------
''' <param><c>fullPathFileName</c> - Full path to the target file</param>
''' --------------------------
''' </summary>
Public Function fileLastModifiedDate(ByVal fullPathFileName As String) As Variant
    
    Dim fso, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(fullPathFileName)

    fileLastModifiedDate = file.DateLastModified
End Function

''' <summary>
''' --------------------------
''' Function <c>fileLastAccessedDate</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Get the last access date for the file
''' --------------------------
'''<returns>Date of the last access</returns>
''' --------------------------
''' <param><c>fullPathFileName</c> - Full path to the target file</param>
''' --------------------------
''' </summary>
Public Function fileLastAccessedDate(ByVal fullPathFileName As String) As Variant
    
    Dim fso, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(fullPathFileName)

    fileLastAccessedDate = file.DateLastAccessed
End Function

''' <summary>
''' --------------------------
''' Function <c>loadFileToString</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Put content of the file to String
''' --------------------------
'''<returns>Content of the target file</returns>
''' --------------------------
''' <param><c>fullPathFileName</c> - Full path to the target file</param>
''' --------------------------
''' </summary>
Public Function loadFileToString(ByVal fullPathFileName As String) As String
Dim iFile As Integer
iFile = FreeFile

Open fullPathFileName For Input As #iFile
loadFileToString = Input(LOF(iFile), iFile)
Close #iFile

End Function

''' <summary>
''' --------------------------
''' Sub <c>saveStringToFile</c>
''' --------------------------
''' version 0.1 (2020-08-28)
''' --------------------------
''' Create file which contains value of the String variable
''' --------------------------
''' <param><c>strContent</c> - Source String variable</param>
''' <param><c>fullPathFileName</c> - Full path to the target file</param>
''' --------------------------
''' </summary>
Public Sub saveStringToFile(ByVal strContent As String, ByVal fullPathFileName As String)
    Dim iFile As Integer
    iFile = FreeFile
    
    Open fullPathFileName For Output As #iFile
    Print #iFile, strContent
    Close #iFile
End Sub
