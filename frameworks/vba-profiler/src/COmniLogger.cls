VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COmniLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfProfiler")
Option Explicit

Public Enum LOG_STREAM_TYPE
    logDebug
    logTXT
    logCSV
    logRange
End Enum

Private mLogTargetType As LOG_STREAM_TYPE
Private mLogTarget As String
Private mLogPosition As Integer
Private mLogHeader As String
Private mLogDelimiter As String
Private mFileID As Integer
Private mIsAppendMode As Boolean

Public Property Get isAppendMode() As Boolean
    isAppendMode = mIsAppendMode

End Property

Public Property Get delimiter() As String
    delimiter = mLogDelimiter

End Property

Public Sub closeLog()
    On Error Resume Next
    
    Select Case mLogTargetType
    Case logTXT, logCSV
        Close #mFileID

    Case Else

    End Select

End Sub

Public Sub initLog(ByVal targetType As LOG_STREAM_TYPE, Optional target As String, Optional newHeader As String = "", Optional delimiter As String = vbTab) ', Optional cols As Integer
    mLogTarget = IIf(targetType = logDebug, "", target)
    mLogTargetType = targetType
    mLogDelimiter = delimiter
    mLogHeader = newHeader
    mFileID = 0
    mLogPosition = 1

End Sub

Public Property Let currentRowInLog(ByVal newPos As Integer)
    mLogPosition = newPos

End Property

Public Property Get currentRowInLog() As Integer
    currentRowInLog = mLogPosition

End Property

Public Sub clear()
    Dim cols As Integer
    
    On Error Resume Next

    Select Case mLogTargetType
    Case logDebug
        Debug.Print vbLf
    
    Case logCSV, logTXT
        Kill mLogTarget

    Case logRange
        enableFastCode True
            cols = lastColumnInTheRow(Range(mLogTarget).Cells(1, 1).row)
            Range(Range(mLogTarget).Cells(1, 1), Range(mLogTarget).Cells(mLogPosition, cols)).clear
        enableFastCode False
    
    Case Else
    
    End Select

    mLogPosition = 1
End Sub

Public Sub openLog()
    mLogPosition = 1
    mIsAppendMode = False
   
    Select Case mLogTargetType
    Case logTXT, logCSV
        mFileID = FreeFile
        Open mLogTarget For Output As #mFileID
    
    Case logRange
        clear
    
    Case Else
    
    End Select

End Sub

Public Sub appendLog()
    mFileID = FreeFile
    mIsAppendMode = True
    
    Select Case mLogTargetType
    Case logTXT, logCSV
        If fileExists(mLogTarget) Then
            Open mLogTarget For Append As #mFileID
        Else
            openLog
        End If
    
    Case logRange
        If Range(mLogTarget).Cells(1, 1) = "" Then
            openLog
            
        Else
            mLogPosition = lastRowInColumn(columnNameByIndex(Range(mLogTarget).Cells(1, 1).Column), Range(mLogTarget).parent.name) - Range(mLogTarget).Cells(1, 1).row + 2
        
        End If
    
    Case Else
    
    End Select

End Sub

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

Public Sub log(ParamArray data() As Variant)
    Dim i As Integer
    Dim toPrint As String
    Dim vArg As Variant
   
    If mLogTargetType <> logRange Then
        toPrint = ""
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

End Sub
