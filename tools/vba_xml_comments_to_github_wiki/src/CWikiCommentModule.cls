VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWikiCommentModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfXMLComments")
''' <summary>
''' --------------------------
''' Module <c>CWikiCommentModule.cls</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' description
''' --------------------------
''' <references>
''' <c>Module/Class name</c>
''' <c>MSugar.bas</c>
''' <c>MText.bas</c>
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

Private iName           As String
Private iVersion        As String
Private iDescription    As String
Private iProceduresList As String
Private iConsts         As String
Private iLocalVariables As String
Private iLicense        As String
Private iReferences     As String
Private iCreated        As String
Private iAuthor         As String


''' <summary>
''' --------------------------
''' Sub <c>loadFromString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Load Module XML comments from String variable
''' --------------------------
''' <param><c>commentStrings</c> - XML comment lines stored in String variable</param>
''' --------------------------
''' </summary>
Public Sub loadFromString(ByVal commentStrings As String)
    
'    'Template for parsing Module:
'    ''' <summary>
'    ''' --------------------------
'    ''' Module <c>Name</c>
'    ''' --------------------------
'    ''' version #.# (2020-08-28)
'    ''' --------------------------
'    ''' description
'    ''' --------------------------
'    ''' <references>
'    ''' <c>Module/Class name</c>
'    ''' </references>
'    ''' --------------------------
'    ''' created 2020-08-28
'    ''' by Sergey Frolov (pet-projects@sergey-frolov.ru)
'    ''' --------------------------
'    ''' </summary>
'    '''
'    ''' <license>
'    ''' This program is free software: you can redistribute it and/or modify
'    ''' it under the terms of the GNU General Public License as published by
'    ''' the Free Software Foundation, either version 3 of the License, or
'    ''' (at your option) any later version.
'    '''
'    ''' This program is distributed in the hope that it will be useful,
'    ''' but WITHOUT ANY WARRANTY; without even the implied warranty of
'    ''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    ''' GNU General Public License for more details.
'    '''
'    ''' You should have received a copy of the GNU General Public License
'    ''' along with this program.  If not, see
'    ''' https://www.gnu.org/licenses/
'    ''' </license>
    
    '---init parser---------------
    Dim arrTokens() As String
    ReDim arrTokens(0)

    'Module name
    addToArray arrTokens, "Module <c>|</c>"
    'Version
    addToArray arrTokens, "version |" & vbCrLf
    'Description
    addToArray arrTokens, "--------------------------" & vbCrLf & "|" & vbCrLf & "--------------------------"
    'References
    addToArray arrTokens, "<references>" & vbCrLf & "|" & vbCrLf & "</references>"
    'Created
    addToArray arrTokens, "created |" & vbCrLf
    'Author
    addToArray arrTokens, "by |" & vbCrLf
    ' --------------------------
    
    Dim i As Integer
    Dim arrResults() As String
    ReDim arrResults(0)
    
    iProceduresList = vbNullString
    iConsts = vbNullString
    iLocalVariables = vbNullString
    iLicense = "Provided under GNU General Public License"
 
    parse commentStrings, arrTokens, arrResults
    
    i = 1
    iName = arrResults(i)
    incr i
    
    iVersion = arrResults(i)
    incr i
    
    iDescription = arrResults(i)
    incr i
    
    iReferences = arrResults(i)
    If iReferences = vbNullString Then
        iReferences = "no other modules required"
    Else
        iReferences = Replace(iReferences, "<c>", "[[")
        iReferences = Replace(iReferences, "</c>", "]]")
    End If
    incr i
    
    iCreated = arrResults(i)
    incr i
    
    iAuthor = arrResults(i)

End Sub


''' <summary>
''' --------------------------
''' Function <c>saveToString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Save Module XML comments to String variable in Github wiki-page format
''' --------------------------
'''<returns>String in Github wiki-page format which contains all Sub XML comments</returns>
''' --------------------------
''' <param><c>none</c></param>
''' --------------------------
''' </summary>
Public Function saveToString() As String
    Dim wikiPage As String

    wikiPage = vbNullString

    incr wikiPage, "## Module `" & iName & "`" & vbCrLf
    
    incr wikiPage, "_**Contains " & iDescription & "**_  " & vbCrLf & vbCrLf
    
    incr wikiPage, "created " & iCreated & "  " & vbCrLf
    
    incr wikiPage, "by " & iAuthor & "  " & vbCrLf
    
    incr wikiPage, "version " & iVersion & "  " & vbCrLf
    
    incr wikiPage, iLicense & "  " & vbCrLf & vbCrLf
    
    incr wikiPage, "#### References: " & "  " & vbCrLf & iReferences & "  " & vbCrLf & vbCrLf
    
    If iLocalVariables <> vbNullString Then incr wikiPage, "#### Local variables:  " & vbCrLf & iLocalVariables & "  " & vbCrLf & vbCrLf
    
    If iConsts <> vbNullString Then incr wikiPage, "#### Constants:  " & vbCrLf & iConsts & vbCrLf & vbCrLf & "----" & vbCrLf
    
    incr wikiPage, "### Contains:  " & vbCrLf & iProceduresList & vbCrLf & "----" & vbCrLf & vbCrLf

    saveToString = wikiPage
End Function

''' <summary>
''' --------------------------
''' Sub <c>addProcedureToList</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Add Sub or Function XML comments to the module list of the procedures
''' --------------------------
''' <param><c>procComment</c> - Sub or Function XML Comments Class</param>
''' --------------------------
''' </summary>
Public Sub addProcedureToList(procComment As IWikiComment)
    
    Dim link As String
    Dim description As String
    
    link = Replace(procComment.name, " ", "-")
    link = Replace(link, "()", "")
    link = Replace(link, "`", "")
    
    If procComment.version <> vbNullString Then
        incr iProceduresList, "* [" & procComment.name & "](#" & link & ") version: " & procComment.version & vbCrLf
    Else
        incr iProceduresList, "* [" & procComment.name & "](#" & link & ")" & vbCrLf
    End If
    
    description = procComment.description
    
'remove example from description
    If InStr(description, "Returns: ") > 0 Then description = Mid(description, 1, InStr(description, "Returns: ") - 1)
    If InStr(description, "<example>") > 0 Then description = Mid(description, 1, InStr(description, "<example>") - 1)
    
    Do While Mid(description, 1, 2) = vbCrLf
        description = Mid(description, 2)
        If Len(description) < 2 Then Exit Do
    Loop
    
    Do While Mid(description, Len(description) - 1, 2) = vbCrLf
        description = Mid(description, 1, Len(description) - 2)
        If Len(description) < 2 Then Exit Do
    Loop
    
    incr iProceduresList, "_" & Trim(description) & "_" & vbCrLf & vbCrLf

End Sub

''' <summary>
''' --------------------------
''' Sub <c>addLocalVariable</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Add locale variable XML comment to the module list of local variables
''' --------------------------
''' <param><c>varComment</c> - String with XML local variable comment</param>
''' --------------------------
''' </summary>
Public Sub addLocalVariable(varComment As String)
    incr iLocalVariables, getTextBetween(varComment, "<param>", "</param>") & vbCrLf
End Sub

''' <summary>
''' --------------------------
''' Sub <c>addConst</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Add const XML comment to the module list of constants
''' --------------------------
''' <param><c>varComment</c> - String with XML const comment</param>
''' --------------------------
''' </summary>
Public Sub addConst(constComment As String)
    incr iConsts, getTextBetween(constComment, "<const>", "</const>") & vbCrLf
End Sub
