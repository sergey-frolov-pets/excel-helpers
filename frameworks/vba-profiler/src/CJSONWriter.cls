VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CJSONWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfProfiler")
Option Explicit

Private mLog As COmniLogger
Private mCurrentLevel As Integer
Private obj_levels() As Integer

Public Property Get currentLevel() As Integer
    currentLevel = mCurrentLevel
End Property

Public Property Let currentLevel(newVal As Integer)
    mCurrentLevel = newVal
End Property

Private Sub log(ByVal toLog As String)
    mLog.log toLog
End Sub

Public Sub openJSON(ByVal targetFile As String)
    mLog.initLog logTXT, targetFile
    mLog.openLog
End Sub

Public Sub appendToJSON(ByVal targetFile As String)
    
    mLog.initLog logTXT, targetFile
    
    Dim prev
    Dim prevStrings
    
    prev = loadFileToString(targetFile)
    If prev = vbNullString Then
         mLog.openLog
    Else
        If Mid(prev, 1, 1) <> "[" Then
           prevStrings = Split(prev, vbLf)
           prev = "[" & vbLf
           Dim pS As Variant
           For Each pS In prevStrings
                If pS <> "" Then incr prev, vbTab & pS & vbLf
           Next
        End If
        prev = Mid(prev, 1, InStrRev(prev, "}")) & ","
        
        Dim fileID As Integer
        fileID = FreeFile
        Open targetFile For Output As #fileID
        Print #fileID, prev
        Close #fileID
        mLog.appendLog
    End If

End Sub

Public Property Get isAppendMode() As Boolean
    isAppendMode = mLog.isAppendMode
End Property

Public Sub closeJSON()
    mLog.closeLog
End Sub

Public Sub addParam(ByVal name As String, ByVal value As String, Optional isLastPeer As Boolean = False)
    log String(mCurrentLevel, vbTab) & """" & name & """: """ & value & """" & IIf(isLastPeer, "", ",")
End Sub

Public Sub addObject(Optional name As String)
    addToArray obj_levels, mCurrentLevel
    
    If name = "" Then
        log String(mCurrentLevel, vbTab) & "{"
    Else
        log String(mCurrentLevel, vbTab) & """" & name & """: {"
    End If
    
    incr mCurrentLevel
End Sub

Public Sub closeObject(Optional isLastPeer As Boolean = False)
    
    If mCurrentLevel = 0 Then
        log "}"
    Else
        log String(mCurrentLevel - 1, vbTab) & "}" & IIf(isLastPeer, "", ",")
    End If

    If UBound(obj_levels) > 0 Then
        ReDim Preserve obj_levels(UBound(obj_levels) - 1)
        mCurrentLevel = obj_levels(UBound(obj_levels)) + 1
    Else
        mCurrentLevel = 0
    End If
End Sub

Public Sub addArray(Optional name As String)
    addToArray obj_levels, mCurrentLevel
    
    If name = "" Then
        log String(mCurrentLevel, vbTab) & "["
    Else
        log String(mCurrentLevel, vbTab) & """" & name & """: ["
    End If
    
    incr mCurrentLevel
End Sub

Public Sub closeArray(Optional isLastPeer As Boolean = False)
    
    If mCurrentLevel = 0 Then
        log "]"
    Else
        log String(mCurrentLevel - 1, vbTab) & "]" & IIf(isLastPeer, "", ",")
    End If
    
    If UBound(obj_levels) > 0 Then
        ReDim Preserve obj_levels(UBound(obj_levels) - 1)
        mCurrentLevel = obj_levels(UBound(obj_levels)) + 1
    Else
        mCurrentLevel = 0
    End If

End Sub

Private Sub Class_Initialize()
    Set mLog = New COmniLogger
    mCurrentLevel = 0
    ReDim obj_levels(0)
End Sub

Private Sub Class_Terminate()
    Set mLog = Nothing
    mCurrentLevel = 0
    ReDim obj_levels(0)
End Sub

