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


Public Function chooseFileOrFolder(ByVal dialogTitle As String, Optional chooseFolder As Boolean = False, Optional initialFileName As String = vbNullString, Optional extentions As String = vbNullString) As String

    Dim directory As String, fileName As String, sheet As Worksheet, total As Integer, fd As Office.FileDialog

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
            .filters.Clear
            
            If extentions = "macro" Then
                .filters.Add "Excel(*.xlsm)", "*.xlsm"
                .filters.Add "Все файлы", "*.*"

            ElseIf extentions = "xlsx" Then
                .filters.Add "Excel(*.xlsx)", "*.xlsx"
                .filters.Add "Все файлы", "*.*"

            Else
                .filters.Add "Excel(*.xls)", "*.xls"
                .filters.Add "Excel(*.xlsx)", "*.xlsx"
                .filters.Add "Excel(*.xlsm)", "*.xlsm"
                .filters.Add "Все файлы", "*.*"

            End If
        
        End If
        
        If .Show = True Then fileName = .SelectedItems(1)

    End With

    chooseFileOrFolder = fileName
End Function

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

Public Function fileExists(ByVal fullFileName As String) As Boolean
    
    If Len(Dir(fullFileName)) = 0 Then
        fileExists = False
    
    Else
        fileExists = True
    
    End If

End Function

Public Function folderExists(ByVal folderFullPath As String) As Boolean
    
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    folderExists = fso.folderExists(folderFullPath)

    Set fso = Nothing
End Function

Public Function fileCreationDate(ByVal fileFullName As String) As Variant
    
    Dim fso, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.GetFile(fileFullName)

    fileCreationDate = file.DateCreated
End Function


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
