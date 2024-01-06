VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COmniLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("OmniLogger framework")

''' <summary>
''' --------------------------
''' Module <c>COmniLogger.cls</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' object for logging into txt/csv files, Excel range and VBA Immediate window
''' --------------------------
''' <references>
''' <c>MSupport.bas</c>
''' </references>
''' --------------------------
''' created 2023-09-01
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

Public Enum LOG_STREAM_TYPE
    logDebug
    logTXT 'synonym of logCSV
    logCSV 'synonym of logTXT
    logRange
End Enum

Private mLogTargetType As LOG_STREAM_TYPE
Private mLogTarget As String

Private mLogHeader As String
Private mLogDelimiter As String

Private mWithTimestamp As Boolean
Private mTimestampFormat As String

Private mFileID As Integer
Private mLogPosition As Integer
Private mIsAppendMode As Boolean

''' <summary>
''' --------------------------
''' Property <c>isAppendMode</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' (read only) Property to check logging mode(new or append)
''' --------------------------
''' <returns>"new"=False or "append"=True</returns>
''' --------------------------
''' </summary>
Public Property Get isAppendMode() As Boolean
    isAppendMode = mIsAppendMode
End Property

''' <summary>
''' --------------------------
''' Property <c>logTarget</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' (read only) path to log
''' --------------------------
''' <returns>File name for logTXT/logCSV, cell address for logRange and "Immediate window" text for logDebulg</returns>
''' --------------------------
''' </summary>
Public Property Get logTarget() As String
    If mLogTargetType = logDebug Then
        logTarget = "Immediate window"
    Else
        logTarget = mLogTarget
    End If
End Property

''' <summary>
''' --------------------------
''' Property <c>timestampFormat</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Timestamp format, "yyyy-mm-dd hh:nn:ss" by default
''' --------------------------
''' <returns>String with date/time format</returns>
''' --------------------------
''' </summary>
Public Property Get timestampFormat() As String
    timestampFormat = mTimestampFormat
End Property

Public Property Let timestampFormat(ByVal newTimestampFormat As String)
    mTimestampFormat = newTimestampFormat
End Property

''' <summary>
''' --------------------------
''' Property <c>withTimeStamp</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' On/off Now() timestamp before the first value into each line with selected format
''' --------------------------
''' <returns>True - on, False - off</returns>
''' --------------------------
''' </summary>
Public Property Get withTimeStamp() As Boolean
    withTimeStamp = mWithTimestamp
    End Property

Public Property Let withTimeStamp(ByVal newWithTimeStamp As Boolean)
    mWithTimestamp = newWithTimeStamp
End Property

''' <summary>
''' --------------------------
''' Property <c>header</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' set/get current header (can be repeated as many times as necessary with logHeader() sub)
''' --------------------------
''' </summary>
Public Property Get header() As String
    header = mLogHeader
End Property

Public Property Let header(ByVal newHeader As String)
    mLogHeader = newHeader
End Property

''' <summary>
''' --------------------------
''' Property <c>delimiter</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' set/get current delimiter (values separator) for each record line
''' --------------------------
''' </summary>
Public Property Get delimiter() As String
    delimiter = mLogDelimiter
End Property

Public Property Let delimiter(ByVal newDelimiter As String)
    mLogDelimiter = newDelimiter
End Property

''' <summary>
''' --------------------------
''' Sub <c>closeLog</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' closes log-file for options logTXT/logCSV
''' --------------------------
''' </summary>
Public Sub closeLog()
    On Error Resume Next
    
    Select Case mLogTargetType
    Case logTXT, logCSV
        Close #mFileID

    Case Else

    End Select

End Sub

''' <summary>
''' --------------------------
''' Sub <c>initLog</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' creates new logging stream (logDebug, logRange, logTXT/logCSV) with additional options (header,delimiter,timestamp, etc.). IMPORTANT: logTXT and logCSV are synonyms.
''' --------------------------
''' <param><c>targetType</c> - [logDebug] for logging into VBA Immediate Window, [logRange] for selected range into current Excel file or [logTXT]/[logCSV] for text-file</param>
''' <param><c>target</c> - full path for text-file(option [logTXT] or [logCSV]), cell address of starting cell including Sheet name for option logRange(i.e., Sheet1!B2)</param>
''' <param><c>newHeader</c> - header string with delimiters(i.e., "Param Name;Param Value")</param>
''' <param><c>newDelimiter</c> - delimiter (values separator) for each record line, vbTab by default. Do not change it for logRange option!</param>
''' <param><c>newWithTimeStamp</c> - on/off timestamp in the beginning of each record. False(off) by default</param>
''' <param><c>newTimestampFormat</c> - Timestamp format, "yyyy-mm-dd hh:nn:ss" by default</param>
''' --------------------------
''' </summary>
Public Sub initLog(ByVal targetType As LOG_STREAM_TYPE, Optional target As String, Optional newHeader As String = "", Optional newDelimiter As String = vbTab, Optional newWithTimeStamp As Boolean = False, Optional newTimestampFormat As String = "yyyy-mm-dd hh:nn:ss")
    mLogTarget = IIf(targetType = logDebug, "", target)
    mLogTargetType = targetType
    mLogDelimiter = newDelimiter
    mLogHeader = newHeader
    mFileID = 0
    mLogPosition = 1
    mWithTimestamp = newWithTimeStamp
    mTimestampFormat = newTimestampFormat
    
End Sub

''' <summary>
''' --------------------------
''' Property <c>currentRowInLog</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' current line counter - for option logRange (ONLY!) can be changed to continue logging from another row, counting records for all other options
''' --------------------------
''' </summary>
Public Property Let currentRowInLog(ByVal newPos As Integer)
    mLogPosition = newPos
End Property

Public Property Get currentRowInLog() As Integer
    currentRowInLog = mLogPosition
End Property

''' <summary>
''' --------------------------
''' Sub <c>clear</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' for logTXT/logCSV and logRange options deletes previous records for logDebug option add empty line to Immediate window
''' --------------------------
''' </summary>
Public Sub clear()
    Dim cols As Integer
    Dim rows As Integer
    
    On Error Resume Next

    Select Case mLogTargetType
    Case logDebug
        Debug.Print vbLf
    
    Case logCSV, logTXT
        Kill mLogTarget

    Case logRange
        cols = lastColumnInTheRow(Range(mLogTarget).Cells(1, 1).row, Range(mLogTarget).Parent.name)
        rows = lastRowInColumn(columnNameByIndex(Range(mLogTarget).Cells(1, 1).Column), Range(mLogTarget).Parent.name)
        Sheets(Range(mLogTarget).Parent.name).Range(Range(mLogTarget).Cells(1, 1).Address, Range(mLogTarget).Cells(rows, cols).Address).clear
    
    Case Else
    
    End Select
 
    mLogPosition = 1
    
End Sub

''' <summary>
''' --------------------------
''' Sub <c>openLog</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' for logTXT/logCSV and logRange options starts log from beginning (deletes previous records and starts from record 1)
''' --------------------------
''' </summary>
Public Sub openLog()
    mIsAppendMode = False
   
    Select Case mLogTargetType
    Case logTXT, logCSV
        mLogPosition = 1
        mFileID = FreeFile
        Open mLogTarget For Output As #mFileID
    
    Case logRange
        mLogPosition = lastRowInColumn(columnNameByIndex(Range(mLogTarget).Cells(1, 1).Column), Range(mLogTarget).Parent.name) - Range(mLogTarget).Cells(1, 1).row + 2
        clear
    
    Case Else
    
    End Select
    
End Sub

''' <summary>
''' --------------------------
''' Sub <c>appendLog</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' for logTXT/logCSV and logRange options continues logging from the last position
''' --------------------------
''' </summary>
Public Sub appendLog()
    
    mIsAppendMode = True
    
    Select Case mLogTargetType
    Case logTXT, logCSV
        If FileExists(mLogTarget) Then
            mFileID = FreeFile
            Open mLogTarget For Append As #mFileID
        Else
            openLog
        End If
    
    Case logRange
        If Range(mLogTarget).Cells(1, 1) = "" Then
            openLog
            
        Else
            mLogPosition = lastRowInColumn(columnNameByIndex(Range(mLogTarget).Cells(1, 1).Column), Range(mLogTarget).Parent.name) - Range(mLogTarget).Cells(1, 1).row + 2
        
        End If
    
    Case Else
    
    End Select

End Sub

''' <summary>
''' --------------------------
''' Sub <c>logHeader</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' records current log header (you can setup header using initLog() sub or change it via header() property)
''' --------------------------
''' </summary>
Public Sub logHeader()

    Select Case mLogTargetType
    Case logDebug
        Debug.Print mLogHeader
    
    Case logCSV, logTXT
        Print #mFileID, mLogHeader
    
    Case logRange
        Dim i As Integer
        i = 1
        
        Dim vArg As Variant
        For Each vArg In Split(mLogHeader, mLogDelimiter)
            Range(mLogTarget).Cells(mLogPosition, i) = vArg
            incr i
        Next vArg
        incr mLogPosition
    
    Case Else
    
    End Select

End Sub

''' <summary>
''' --------------------------
''' Sub <c>log</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' records new log line - it can be one Value (.log Value_1), or ParamArray of Values (.log Value_1, Value_2, …, Value_N)
''' --------------------------
''' <param><c>data</c> - ParamArray of values for logging (e.g. log "A",1,"B",2 )</param>
''' --------------------------
''' </summary>
Public Sub log(ParamArray data() As Variant)
    Dim i As Integer
    Dim toPrint As String
    Dim vArg As Variant
   
    If mLogTargetType <> logRange Then
        If mWithTimestamp Then
            toPrint = Format(Now(), mTimestampFormat)
        Else
            toPrint = ""
        End If
        
        For Each vArg In data
            toPrint = toPrint & IIf(toPrint = "", vArg, mLogDelimiter & vArg)
        Next vArg
    
    End If

    Select Case mLogTargetType
    Case logDebug
        Debug.Print toPrint
    
    Case logCSV, logTXT
        Print #mFileID, toPrint

    Case logRange
        i = 1
        If mWithTimestamp Then
            Range(mLogTarget).Cells(mLogPosition, i).NumberFormat = mTimestampFormat
            Range(mLogTarget).Cells(mLogPosition, i) = Format(Now(), mTimestampFormat)
            incr i
        End If
        
        If UBound(data) > 1 Then
            For Each vArg In data
                Range(mLogTarget).Cells(mLogPosition, i) = vArg
                incr i
            Next vArg
        Else
            Dim dat
            dat = Split(data(0), mLogDelimiter)
            For Each vArg In dat
                Range(mLogTarget).Cells(mLogPosition, i) = vArg
                incr i
            Next vArg
        
        End If
        incr mLogPosition
    
    Case Else
    
    End Select

End Sub

''' <summary>
''' --------------------------
''' Sub <c>logArray</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' records new log line from values array
''' --------------------------
''' <param><c>data</c> - array of values for logging</param>
''' --------------------------
''' </summary>
Public Sub logArray(data As Variant)
    Dim i As Integer
    Dim toPrint As String
    Dim vArg As Variant
   
    If mLogTargetType <> logRange Then
        If mWithTimestamp Then
            toPrint = Format(Now(), mTimestampFormat)
        Else
            toPrint = ""
        End If
        
        For Each vArg In data
            toPrint = toPrint & IIf(toPrint = "", vArg, mLogDelimiter & vArg)
        Next vArg
    
    End If

    Select Case mLogTargetType
    Case logDebug
        Debug.Print toPrint
    
    Case logCSV, logTXT
        Print #mFileID, toPrint

    Case logRange
        i = 1
        If mWithTimestamp Then
            Range(mLogTarget).Cells(mLogPosition, i).NumberFormat = mTimestampFormat
            Range(mLogTarget).Cells(mLogPosition, i) = Format(Now(), mTimestampFormat)
            incr i
        End If
        
        If UBound(data) > 1 Then
            For Each vArg In data
                Range(mLogTarget).Cells(mLogPosition, i) = vArg
                incr i
            Next vArg
        Else
            Dim dat
            dat = Split(data(0), mLogDelimiter)
            For Each vArg In dat
                Range(mLogTarget).Cells(mLogPosition, i) = vArg
                incr i
            Next vArg
        
        End If
        incr mLogPosition
    
    Case Else
    
    End Select

End Sub

Private Sub Class_Initialize()
    mLogTarget = ""
    mLogTargetType = logDebug
    mLogDelimiter = ""
    mLogHeader = ""
    mFileID = 0
    mLogPosition = 1
    mIsAppendMode = False
    mWithTimestamp = False
    mTimestampFormat = "yyyy-mm-dd hh:nn:ss"
    
End Sub

