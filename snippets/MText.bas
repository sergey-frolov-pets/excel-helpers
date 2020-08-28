Attribute VB_Name = "MText"
'@Folder("sfSnippets")
''' <summary>
''' --------------------------
''' Module <c>MText.bas</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' procedures to work with text stored in String variables and/or arrays
''' --------------------------
''' <references>
'''   <c>none</c>
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
''' Function <c>hasOneOfKeywords</c>
''' --------------------------
''' Check if text contains one of the keywords stored in array
''' --------------------------
''' <returns>True if text contains ANY of the keywords and False if not. Function can be inverted</returns>
''' --------------------------
''' <param><c>text</c> - Source text</param>
''' <param><c>arrKeyWords</c> - Array with keywords</param>
''' <param><c>inversion</c> - Optional. True if we have to invert function result</param>
''' --------------------------
''' created 2020-08-20
''' --------------------------
''' </summary>
Public Function hasOneOfKeywords(ByVal text As String, arrKeyWords As Variant, Optional inversion As Boolean = False) As Boolean
    Dim i As Integer

    For i = 1 To UBound(arrKeyWords)
        If InStr(LCase(text), LCase(arrKeyWords(i))) > 0 Then
            If inversion Then
                hasOneOfKeywords = False
            
            Else
                hasOneOfKeywords = True
            
            End If
            
            Exit Function
        
        End If
    
    Next

    If inversion Then
        hasOneOfKeywords = True
    
    Else
        hasOneOfKeywords = False
    
    End If

End Function

Public Function getTextAfter(ByVal fromText As String, ByVal afterText As String) As String
    
    If InStr(fromText, afterText) = 0 Then
        getTextAfter = vbNullString
    
    Else
        getTextAfter = Mid(fromText, InStr(fromText, afterText) + Len(afterText))
    
    End If

End Function

Public Function getTextBetween(ByVal inText As String, ByVal startMarker As String, ByVal endMarker As String, Optional fromPosition As Integer = 1) As String
    Dim startPos As Long, endPos As Long, curPos As Long

    If fromPosition > 0 Then curPos = fromPosition Else curPos = 1

    startPos = InStr(curPos, inText, startMarker, False) + Len(startMarker)

    If startPos < Len(startMarker) + 1 Then
        getTextBetween = "<NOT FOUND>"
    
    Else
        endPos = InStr(startPos, inText, endMarker, False) - 1
        If endPos < 1 Then
            getTextBetween = Mid(inText, startPos)
            curPos = Len(inText)
        
        Else
            getTextBetween = Mid(inText, startPos, endPos - startPos + 1)
            curPos = endPos + 1
        
        End If
    
    End If
End Function

Public Function parse(ByVal textToParse As String, arrTokens() As String, arrResults() As String) As Integer
    Dim i As Integer
    Dim foundValues As Integer
    Dim nextPos As Long
    Dim curTextToParse As String
    Dim curRes As String
    Dim nextStartPos As Long

    Dim token

    foundValues = 0

    ReDim arrResults(UBound(arrTokens))

    curTextToParse = textToParse

    For i = 1 To UBound(arrTokens)
        token = Split(arrTokens(i), "|")
        curRes = getTextBetween(curTextToParse, token(0), token(1))
        
        If curRes <> "<NOT FOUND>" Then
            arrResults(i) = curRes
            incr foundValues
        End If
    
        nextStartPos = InStr(curTextToParse, token(1))
        
        If nextStartPos > Len(curTextToParse) Then Exit For
        
        curTextToParse = Mid(curTextToParse, nextStartPos)
    Next

    parse = foundValues
End Function

Public Function parseRecords(ByVal textToParse As String, arrRecordTokens() As String, arrResults() As String) As Integer
    
    Dim firstRecordToken As String
    Dim startPos As Long
    Dim curText As String
    Dim curRec() As String
    Dim tokensCount As Integer
    
    
    tokensCount = UBound(arrRecordTokens)
    recordsCount = 0
    ReDim arrResults(tokensCount, recordsCount)
    
    firstRecordToken = Split(arrRecordTokens(1), "|")(0) 'берем открывающую часть первого токена
    lenOfFirstRecordToken = Len(firstRecordToken)
    
    startPos = InStr(textToParse, firstRecordToken)
   
    Do While startPos > 0
        
        endOfCurrentRecord = InStr(startPos + lenOfFirstRecordToken, textToParse, firstRecordToken) - 1
        If endOfCurrentRecord = -1 Then endOfCurrentRecord = Len(textToParse)
        
        If parse(Mid(textToParse, startPos, endOfCurrentRecord), arrRecordTokens, curRec) > 0 Then
           incr recordsCount
           ReDim Preserve arrResults(tokensCount, recordsCount)
           For i = 1 To tokensCount
               arrResults(i, recordsCount) = curRec(i)
           Next
        End If
        
        If endOfCurrentRecord < Len(textToParse) Then
            startPos = endOfCurrentRecord + 1
        Else
            startPos = 0
        End If
    Loop

parseRecords = recordsCount
End Function
