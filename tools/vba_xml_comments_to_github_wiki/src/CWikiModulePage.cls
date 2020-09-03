VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWikiModulePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfXMLComments")
''' <summary>
''' --------------------------
''' Module <c>CWikiModulePage.cls</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Class of Module XML Comments (for any of VBA files - *.bas, *.cls, *.frm)
''' --------------------------
''' <references>
''' <c>IWikiComment.cls</c>
''' <c>CWikiCommentModule.cls</c>
''' <c>MSugar.bas</c>
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

''' <param><c>iComments()</c> - Array of extracted XML-comments for VBA-file</param>
Dim iComments() As IWikiComment
''' <param><c>iModule</c> - XML-comment for VBA-file</param>
Dim iModule  As CWikiCommentModule

Public Property Get countComments() As Integer
    countComments = UBound(iComments)
End Property

Public Sub reset()
    ReDim iComments(0)
    Set iModule = Nothing
End Sub

Private Sub Class_Initialize()
    reset
End Sub


''' <summary>
''' --------------------------
''' Function <c>commentType</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Define comment type by XML comment content
''' --------------------------
'''<returns>comment type</returns>
''' --------------------------
''' <param><c>commentStrings</c> - XML comment lines stored in String variable</param>
''' --------------------------
''' </summary>
Public Function commentType(ByVal commentStrings As String) As String
    If InStr(1, Trim(commentStrings), "Module <c>") > 0 Then
        commentType = "module"

    ElseIf InStr(1, Trim(commentStrings), "Function <c>") > 0 Then
        commentType = "function"

    ElseIf InStr(1, Trim(commentStrings), "Sub <c>") > 0 Then
        commentType = "sub"

    ElseIf InStr(1, Trim(commentStrings), "<const><c>") > 0 Then
        commentType = "const"

    ElseIf InStr(1, Trim(commentStrings), "<param><c>") > 0 Then
        commentType = "param"

    Else
        commentType = "unknown"

    End If

End Function

''' <summary>
''' --------------------------
''' Sub <c>addCommentFromString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Add XML comment to iComments() array from String variable
''' --------------------------
''' <param><c>commentStrings</c> - XML comment lines stored in String variable</param>
''' --------------------------
''' </summary>
Public Sub addCommentFromString(ByVal commentStrings As String)
    Dim iComment As IWikiComment

    Select Case commentType(commentStrings)
    Case "module":
        If iModule Is Nothing Then
            Set iModule = New CWikiCommentModule
            iModule.loadFromString commentStrings
       
        End If
            
    Case "function":
        Set iComment = New CWikiCommentFunction
        iComment.loadFromString commentStrings
        If Not iModule Is Nothing Then iModule.addProcedureToList iComment
        addToArray iComments, iComment
       
    Case "sub":
        Set iComment = New CWikiCommentSub
        iComment.loadFromString commentStrings
        If Not iModule Is Nothing Then iModule.addProcedureToList iComment
        addToArray iComments, iComment
       
    Case "const":
        If Not iModule Is Nothing Then iModule.addConst commentStrings
       
    Case "param":
        If Not iModule Is Nothing Then iModule.addLocalVariable commentStrings
       
    Case Else:
        MsgBox "Wrong format!" & vbCrLf & commentStrings, vbExclamation
            
    End Select

End Sub

''' <summary>
''' --------------------------
''' Function <c>saveToString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Save XML comments from iComments() array to String variable in Github wiki-page format
''' --------------------------
'''<returns>String in Github wiki-page format which contains all comments from iComments() array (one VBA-file)</returns>
''' --------------------------
''' <param><c>none</c></param>
''' --------------------------
''' </summary>
Public Function saveToString() As String
    Dim tmpWiki As String
    Dim comment As Variant
    Dim i       As Integer

    If Not iModule Is Nothing Then tmpWiki = iModule.saveToString()

    For i = 1 To UBound(iComments)
        incr tmpWiki, vbCrLf & iComments(i).saveToString()

    Next
    
    tmpWiki = Replace(tmpWiki, "<param>", "")
    tmpWiki = Replace(tmpWiki, "</param>", vbCrLf)
    
    tmpWiki = Replace(tmpWiki, "<const>", "")
    tmpWiki = Replace(tmpWiki, "</const>", "")
    
    tmpWiki = Replace(tmpWiki, "<c>", "`")
    tmpWiki = Replace(tmpWiki, "</c>", "`")

    tmpWiki = Replace(tmpWiki, "<example>", vbCrLf)
    tmpWiki = Replace(tmpWiki, "</example>", vbCrLf)
    
    saveToString = tmpWiki
End Function
