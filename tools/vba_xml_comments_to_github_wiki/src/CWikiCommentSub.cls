VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWikiCommentSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfXMLComments")
''' <summary>
''' --------------------------
''' Module <c>CWikiCommentSub.cls</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Class of Sub XML Comments (for any of VBA files - *.bas, *.cls, *.frm)
''' --------------------------
''' <references>
''' <c>IWikiComment.cls</c>
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


Implements IWikiComment

Private iName           As String
Private iVersion        As String
Private iDescription    As String
Private iParams         As String


''' <summary>
''' --------------------------
''' Sub <c>IWikiComment_loadFromString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Load Sub XML comments from String variable
''' --------------------------
''' <param><c>commentStrings</c> - XML comment lines stored in String variable</param>
''' --------------------------
''' </summary>
Public Sub IWikiComment_loadFromString(ByVal commentStrings As String)
    
' 'Template for parsing Sub():
' ''' <summary>
' ''' --------------------------
' ''' Sub <c>Name</c>
' ''' --------------------------
' ''' version #.# (2020-08-28)
' ''' --------------------------
' ''' Description
' ''' --------------------------
' ''' <param><c>Name</c> - Description</param>
' ''' --------------------------
' ''' </summary>
    
'---init parser---------------
    Dim arrTokens() As String
    ReDim arrTokens(0)
    
' Sub name
    addToArray arrTokens, "Sub <c>|</c>"
' Version
    If InStr(commentStrings, "version ") > 0 Then addToArray arrTokens, "version |" & vbCrLf
' Description
    addToArray arrTokens, "--------------------------" & vbCrLf & "|" & vbCrLf & "--------------------------"
' Parameters
    addToArray arrTokens, "<param>|" & vbCrLf & "--------------------------"
' --------------------------

    Dim i As Integer
    Dim arrResults() As String
    ReDim arrResults(0)
 
    parse commentStrings, arrTokens, arrResults
    
    i = 1
    iName = arrResults(i)
    incr i
    
    If InStr(commentStrings, "version ") > 0 Then
        iVersion = arrResults(i)
        incr i
    End If
    
    iDescription = arrResults(i)
    incr i
    
    iParams = arrResults(i)
End Sub


''' <summary>
''' --------------------------
''' Function <c>saveToString</c>
''' --------------------------
''' version 0.1 (2020-09-02)
''' --------------------------
''' Save Sub XML comments to String variable in Github wiki-page format
''' --------------------------
'''<returns>String in Github wiki-page format which contains all Sub XML comments</returns>
''' --------------------------
''' <param><c>none</c></param>
''' --------------------------
''' </summary>
Public Function IWikiComment_saveToString() As String
    Dim wikiPage As String

    wikiPage = vbNullString

    incr wikiPage, "### Sub `" & iName & "()`  " & vbCrLf
    
    incr wikiPage, "_" & iDescription & "_" & vbCrLf & vbCrLf
    
    If iVersion <> vbNullString Then incr wikiPage, "version " & iVersion & "  " & vbCrLf
    
    If iParams <> vbNullString Then incr wikiPage, "#### Parameters:  " & vbCrLf & iParams & vbCrLf

    incr wikiPage, vbCrLf & "----" & vbCrLf
    
    IWikiComment_saveToString = wikiPage

End Function

Private Function IWikiComment_name() As String
IWikiComment_name = "Sub `" & iName & "()`"

End Function

Private Function IWikiComment_description() As String
IWikiComment_description = iDescription

End Function

Private Function IWikiComment_version() As String
IWikiComment_version = iVersion

End Function
