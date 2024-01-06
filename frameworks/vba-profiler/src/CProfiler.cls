VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("sfProfiler")

Option Explicit

Private Type PROC_CALLS_INFO
    procName As String
    procComments As String
    procStart As Double
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private mStartSessionTime As SYSTEMTIME
Private mStartProcedureTickTime As Double
Private mStopSessionTime As SYSTEMTIME
Private mDurationSession As Double

Private mIsSessionStarted As Boolean
Private mLastProcedure As String

Private mRuns As Dictionary
Private mSession() As PROC_CALLS_INFO
Private mCallsTree() As PROC_CALLS_INFO
Private mActiveCalls As Integer

Private mTabHeader() As String

Private mLog As COmniLogger

Public Sub startSession()
    
    If Not (mRuns Is Nothing) Then mRuns.RemoveAll
    Set mRuns = New Dictionary
    
    mActiveCalls = 0
    ReDim mCallsTree(mActiveCalls)
    
    ReDim mSession(0)
    mLastProcedure = ""
    
    mIsSessionStarted = True
    mDurationSession = GetTickCount
    GetLocalTime mStartSessionTime

End Sub

Public Sub stopSession()
    Dim nowIs As Double, i As Integer
    
    mIsSessionStarted = False
    nowIs = GetTickCount
    
    If mStartProcedureTickTime > 0 Then
        For i = UBound(mCallsTree) To 1 Step -1
            mRuns(mCallsTree(i).procName) = mRuns(mCallsTree(i).procName) & ";" & nowIs - mCallsTree(i).procStart
        Next
                
    End If
    
    mLastProcedure = ""
    ReDim mCallsTree(0)
    mStartProcedureTickTime = 0
    
    mDurationSession = nowIs - mDurationSession
    GetLocalTime mStopSessionTime

End Sub

Private Sub stat(Optional forProcedure As String)
    If forProcedure = "" Then
        sessionStat
    Else
        sessionStatOnProcedure forProcedure
    End If

End Sub

Private Sub logCallsTree()
    Dim i As Integer, cap As String

    cap = "CALLS TREE:"
    log vbLf & cap & vbLf & String(Len(cap), "-")

    For i = 1 To UBound(mSession)
        With mSession(i)
            log String(.procStart - 1, vbTab) & .procName & IIf(.procComments = "", "", vbTab & "// " & .procComments)
        End With
    Next

End Sub

Private Sub toJSONProcedure(ByVal forProcedure As String, jsonObj As CJSONWriter, Optional isLastProcedure As Boolean = False)
    Dim runs() As String
    
    Dim comments() As String
    Dim orders() As String
    Dim duration() As Double
    
    Dim proc_total_duration As Double
    Dim j As Integer, h As Integer, g As Integer

    With jsonObj
        .addObject forProcedure
        If Not mRuns.exists(forProcedure) Then
            .addParam "error", "procedure <<" & forProcedure & ">> is not found"
            .closeObject isLastProcedure
            Exit Sub
        End If

        runs = Split(mRuns(forProcedure), ";")
    
        ReDim duration(UBound(runs))
        ReDim comments(UBound(runs))
        ReDim orders(UBound(runs))
        
        g = 0
        For h = 1 To UBound(mSession)
            With mSession(h)
                If .procName = forProcedure Then
                    incr g
                    comments(g) = .procComments
                    orders(g) = h
                End If
            End With
        Next
    
        proc_total_duration = 0
        For j = 1 To UBound(runs)
            duration(j) = Val(runs(j))
            incr proc_total_duration, duration(j)
        Next
        RemoveFromArray duration, 0
    
        Dim countRuns As Integer, minTime As Double, maxTime As Double, avgTime As Double
        countRuns = UBound(runs)
        minTime = Application.Min(duration)
        maxTime = Application.Max(duration)
        avgTime = Application.Average(duration)
    
        .addParam "total runs", countRuns
        .addParam "total runtime", msToTimeString(proc_total_duration)
        .addParam "share of session", Round(100 * proc_total_duration / mDurationSession, 0) & "%"
        .addParam "min runtime", msToTimeString(minTime)
        .addParam "max runtime", msToTimeString(maxTime)
        .addParam "avg runtime", msToTimeString(avgTime)
     
        .addArray "runs"
        For j = 1 To UBound(runs)
            .addObject
            .addParam "run", j
            .addParam "order in session", orders(j)
            .addParam "duration", msToTimeString(duration(j - 1))
            .addParam "share of session", Round(100 * duration(j - 1) / mDurationSession, 0) & "%", (comments(j) = "")
            If comments(j) <> "" Then .addParam "comments", comments(j), True
            .closeObject (j = UBound(runs))
        Next
        .closeArray True
        .closeObject isLastProcedure
    End With
End Sub

Public Sub toTXT(ByVal toFile As String)
    mLog.initLog logTXT, toFile
    mLog.openLog
    
    Dim i As Integer, caption As String
    
    stat
    logCallsTree
    
    log ""
    caption = "PROCEDURES:"
    
    With mStartSessionTime
        log caption
        log String(Len(caption), "-")
    
    End With
    
    Dim k As String
    For i = 1 To mRuns.count
        caption = "PROCEDURE <<" & mRuns.keys(i - 1) & ">> INFO:"
        log caption
        log String(Len(caption), "-")
        
        k = mRuns.keys(i - 1)
        stat k
    
    Next
        
    mLog.closeLog

End Sub

Private Sub initTabHeader()
    ReDim mTabHeader(0)
    addToArray mTabHeader, "Session start"
    addToArray mTabHeader, "Session end"
    addToArray mTabHeader, "Procedure"
    addToArray mTabHeader, "Run"
    addToArray mTabHeader, "Level"
    addToArray mTabHeader, "Order"
    addToArray mTabHeader, "Duration,ms"
    addToArray mTabHeader, "Share,%"
    addToArray mTabHeader, "Total runs,ms"
    addToArray mTabHeader, "Total,ms"
    addToArray mTabHeader, "Total Share,%"
    addToArray mTabHeader, "Average,ms"
    addToArray mTabHeader, "Min,ms"
    addToArray mTabHeader, "Max,ms"
    addToArray mTabHeader, "Comments"

End Sub

Public Sub toCSV(ByVal toFile As String, Optional appendToEnd As Boolean = False)
    logTabular logCSV, toFile, appendToEnd

End Sub

Private Sub logTabular(ByVal targetType As LOG_STREAM_TYPE, Optional target As String, Optional appendToEnd As Boolean = False)
    initTabHeader
    
    mLog.initLog targetType, target, Mid(Join(mTabHeader, mLog.delimiter), Len(mLog.delimiter) + 1)
    
    If appendToEnd Then mLog.appendLog Else mLog.openLog
        
    If Not mLog.isAppendMode Then mLog.logHeader
    
    sessionStatTabular
    
    Dim i As Integer
    For i = 1 To mRuns.count
        sessionStatOnProcedureTabular mRuns.keys(i - 1)
    Next
    
    mLog.closeLog
End Sub

Private Sub toDebugTabular()
    logTabular logDebug

End Sub

Public Sub toRange(ByVal startCellAddress As String, Optional appendToEnd As Boolean = False)
    logTabular logRange, startCellAddress, appendToEnd

End Sub

Public Sub toDebug(Optional inTabularForm As Boolean = False)
    
    If inTabularForm Then
        toDebugTabular
    
    Else
        mLog.initLog logDebug
        
        Dim i As Integer, caption As String
        
        stat
        logCallsTree
    
        log ""
        caption = "PROCEDURES:"
        
        With mStartSessionTime
            log caption
            log String(Len(caption), "-")
        End With
        
        Dim k As String
        For i = 1 To mRuns.count
            caption = "PROCEDURE <<" & mRuns.keys(i - 1) & ">> INFO:"
            log caption
            log String(Len(caption), "-")
            k = mRuns.keys(i - 1)
            stat k
        Next
    
    End If

End Sub

Public Sub toJSON(ByVal toFile As String, Optional appendToEnd As Boolean = False)
    Dim jw As New CJSONWriter
    Dim i As Integer
    Dim listOfRuns As String, runNum As Integer
    Dim openObjCount As Integer
    
    With jw
        If appendToEnd Then
            .appendToJSON toFile
            If .isAppendMode Then .currentLevel = 1
        Else
            .openJSON toFile
            .currentLevel = 0
        End If
    
        .addObject
            .addObject "session"
                .addParam "start", sysTimeToTime(mStartSessionTime)
                .addParam "stop", sysTimeToTime(mStopSessionTime)
                .addParam "duration", msToTimeString(mDurationSession)
        
                '-------------------------------------------------------
                .addObject "calls tree"
            
                listOfRuns = "<"
                
                For i = 1 To UBound(mSession)
                    
                    incr listOfRuns, mSession(i).procName & "><"
                    runNum = countSubStrings(listOfRuns, "<" & mSession(i).procName & ">")
                    
                    .addObject mSession(i).procName & "_run_" & Format(runNum, "00")
                    incr openObjCount
                    
                    If i < UBound(mSession) Then
                        If mSession(i).procStart = mSession(i + 1).procStart Then
                                    
                            If mSession(i).procComments <> "" Then .addParam "comments", mSession(i).procComments, True
                            .closeObject
                            incr openObjCount, -1
                        
                        ElseIf mSession(i).procStart > mSession(i + 1).procStart Then
                            If mSession(i).procComments <> "" Then .addParam "comments", mSession(i).procComments, True
                            .closeObject True
                            incr openObjCount, -1
                            .closeObject
                            incr openObjCount, -1
                        
                        Else
                            If mSession(i).procComments <> "" Then .addParam "comments", mSession(i).procComments
                        
                        End If
                    
                    Else
                        If mSession(i).procComments <> "" Then .addParam "comments", mSession(i).procComments, True
                    
                    End If
                
                Next
        
                For i = 1 To openObjCount
                    .closeObject True
                Next
            
                .closeObject
        '-------------------------------------------------------
        
                .addObject "procedures"
        
                For i = 1 To mRuns.count
                    toJSONProcedure mRuns.keys(i - 1), jw, (i = mRuns.count)
                Next
        
           .closeObject True
        .closeObject True
      .closeObject True
        
      If .isAppendMode Then .closeArray True
     
     .closeJSON
    
    End With

End Sub

Private Sub sessionStat()
    Dim runs() As String
    Dim duration() As Double
    Dim proc_total_duration() As Double
    Dim caption As String
    Dim i As Integer
    
    caption = "PROFILER SESSION INFO:"
    
    With mStartSessionTime
        log caption
        log String(Len(caption), "-")
        log "Session started @" & sysTimeToTime(mStartSessionTime)
    
    End With

    ReDim proc_total_duration(mRuns.count)
    
    For i = 1 To mRuns.count
        runs = Split(mRuns.Items(i - 1), ";")
        ReDim duration(UBound(runs))
        
        proc_total_duration(i) = 0
        
        Dim j As Integer
        For j = 1 To UBound(runs)
            duration(j) = Val(runs(j))
            incr proc_total_duration(i), duration(j)
        
        Next
    
        RemoveFromArray duration, 0
        
        Dim countRuns As Integer, minTime As Double, maxTime As Double, avgTime As Double
        countRuns = UBound(runs)
        minTime = Application.Min(duration)
        maxTime = Application.Max(duration)
        avgTime = Application.Average(duration)
    
        log vbTab & fromTemplate("{%1%}: {%2%} run(s) with min {%3%}ms, max {%4%}ms, avg {%5%}ms, total {%6%}ms, share of session {%7%}%", _
                                 mRuns.keys(i - 1), countRuns, Round(minTime, 0), Round(maxTime, 0), Round(avgTime, 0), _
                                 proc_total_duration(i), Round(100 * proc_total_duration(i) / mDurationSession, 0))
    Next
    
    ReDim duration(0)
    ReDim proc_total_duration(0)
    
    log "Session stopped @" & sysTimeToTime(mStopSessionTime)

End Sub

Private Sub sessionStatOnProcedure(ByVal forProcedure As String)
    Dim runs() As String
    Dim comments() As String
    Dim duration() As Double
    Dim proc_total_duration As Double
    Dim j As Integer, h As Integer, g As Integer

    If Not mRuns.exists(forProcedure) Then
        log "!!!"
        log "!!! Procedure <<" & forProcedure & ">> is not found !!!"
        log "!!!"
        log ""
        Exit Sub
    
    End If

   
    log "Runs:"

    runs = Split(mRuns(forProcedure), ";")
    
    ReDim duration(UBound(runs))
    ReDim comments(UBound(runs))
    
    g = 0
    For h = 1 To UBound(mSession)
        With mSession(h)
            If .procName = forProcedure Then
                incr g
                comments(g) = IIf(.procComments = "", "", "// " & .procComments)
            End If
        End With
    Next
    
    proc_total_duration = 0
    For j = 1 To UBound(runs)
        duration(j) = Val(runs(j))
        incr proc_total_duration, duration(j)
        log fromTemplate(vbTab & "run#{%1%}: {%2%}ms share of session {%3%}% {%4%}", j, duration(j), _
            Round(100 * duration(j) / mDurationSession, 0), IIf(comments(j) = "", "", vbTab & comments(j)))
    Next

    RemoveFromArray duration, 0
    
    Dim countRuns As Integer, minTime As Double, maxTime As Double, avgTime As Double
    countRuns = UBound(runs)
    minTime = Application.Min(duration)
    maxTime = Application.Max(duration)
    avgTime = Application.Average(duration)
    
    log fromTemplate("All runs: {%1%} run(s) with min {%2%}ms, max {%3%}ms, avg {%4%}ms, total {%5%}ms, share of session {%6%}%" & vbLf, _
                      countRuns, Round(minTime, 0), Round(maxTime, 0), Round(avgTime, 0), _
                      proc_total_duration, Round(100 * proc_total_duration / mDurationSession, 0))
    
End Sub

Private Sub sessionStatOnProcedureTabular(ByVal forProcedure As String)

    'Session start
    'Session end
    'Procedure
    'Run
    'Level
    'Duration,ms
    'Share,%
    'Total runs,ms
    'Total,ms
    'Total Share,%
    'Average,ms
    'Min,ms
    'Max,ms
    'Comments
        
    Dim runs() As String
    Dim levels() As Integer
    Dim orders() As Integer
    Dim comments() As String
    Dim duration() As Double
    
    Dim proc_total_duration As Double
    Dim i As Integer, j As Integer, h As Integer, g As Integer
    Dim toPrint() As String
  
    If Not mRuns.exists(forProcedure) Then
        ReDim toPrint(4)
        addToArray toPrint, sysTimeToTime(mStartSessionTime)
        addToArray toPrint, sysTimeToTime(mStopSessionTime)
        addToArray toPrint, forProcedure
        addToArray toPrint, "NOT FOUND!!!"
        log Mid(Join(toPrint, mLog.delimiter), Len(mLog.delimiter) + 1)
        Exit Sub
    
    End If

    runs = Split(mRuns(forProcedure), ";")
    
    ReDim duration(UBound(runs))
    ReDim comments(UBound(runs))
    ReDim levels(UBound(runs))
    ReDim orders(UBound(runs))
    
    g = 0
    For h = 1 To UBound(mSession)
        With mSession(h)
            If .procName = forProcedure Then
                incr g
                comments(g) = .procComments
                levels(g) = .procStart
                orders(g) = h
            End If
        End With
    Next
    
    proc_total_duration = 0
    For j = 1 To UBound(runs)
        duration(j) = Val(runs(j))
        incr proc_total_duration, duration(j)
    Next

    RemoveFromArray duration, 0
    
    Dim countRuns As Integer, minTime As Double, maxTime As Double, avgTime As Double
    countRuns = UBound(runs)
    minTime = Application.Min(duration)
    maxTime = Application.Max(duration)
    avgTime = Round(Application.Average(duration), 0)
    
    ReDim toPrint(15)
    
    For j = 1 To UBound(runs)
        i = 0
        'Session start
        incr i
        toPrint(i) = sysTimeToTime(mStartSessionTime)
        'Session end
        incr i
        toPrint(i) = sysTimeToTime(mStopSessionTime)
        'Procedure
        incr i
        toPrint(i) = forProcedure
        'Run
        incr i
        toPrint(i) = j
        
        'Level
        incr i
        toPrint(i) = levels(j)
        
        'Order
        incr i
        toPrint(i) = orders(j)

        'Duration,ms
        incr i
        toPrint(i) = duration(j - 1)

        'Share,%
        incr i
        toPrint(i) = Round(100 * duration(j - 1) / mDurationSession, 0)

        'Total runs
        incr i
        toPrint(i) = countRuns

        'Total duration,ms
        incr i
        toPrint(i) = proc_total_duration

        'Total Share,%
        incr i
        toPrint(i) = Round(100 * proc_total_duration / mDurationSession, 0)

        'Average,ms
        incr i
        toPrint(i) = avgTime
        'Min,ms
        incr i
        toPrint(i) = minTime
        'Max,ms
        incr i
        toPrint(i) = maxTime
        'Comments
        incr i
        toPrint(i) = comments(j)
        
        log Mid(Join(toPrint, mLog.delimiter), Len(mLog.delimiter) + 1)

    Next
    
End Sub

Private Sub sessionStatTabular()

    'Session start
    'Session end
    'Procedure
    'Run
    'Level
    'Order
    'Duration,ms
    'Share,%
    'Comments
    'Total runs,ms
    'Total,ms
    'Total Share,%
    'Average,ms
    'Min,ms
    'Max,ms
    
    Dim i As Integer, toPrint() As String
   
    ReDim toPrint(14)
    
    i = 0
    'Session start
    incr i
    toPrint(i) = sysTimeToTime(mStartSessionTime)
    'Session end
    incr i
    toPrint(i) = sysTimeToTime(mStopSessionTime)
    'Procedure
    incr i
    toPrint(i) = "<<Session>>"
    'Run
    incr i
    toPrint(i) = 0
        
    'Level
    incr i
    toPrint(i) = 0
    'Order
    incr i
    toPrint(i) = 0
        
    'Duration,ms
    incr i
    toPrint(i) = mDurationSession

    'Share,%
    incr i
    toPrint(i) = 100

    'Total runs
    incr i
    toPrint(i) = UBound(mSession)

    'Total duration,ms
    incr i
    toPrint(i) = mDurationSession

    'Total Share,%
    incr i
    toPrint(i) = 100

    'Average,ms
    incr i
    toPrint(i) = mDurationSession
    'Min,ms
    incr i
    toPrint(i) = mDurationSession
    'Max,ms
    incr i
    toPrint(i) = mDurationSession
        
    log Mid(Join(toPrint, mLog.delimiter), Len(mLog.delimiter) + 1)
    
End Sub

Public Sub p_(Optional forProcedure As String, Optional inModule As String, Optional comments As String)
    If mIsSessionStarted Then
        
        Dim nowIs As Double, i As Integer
        
        If (forProcedure = "" And inModule = "") _
        Or mLastProcedure & "/" = inModule & "." & forProcedure _
           Or (mLastProcedure & "/" = forProcedure And inModule = "") Then

            If mStartProcedureTickTime > 0 Then
                nowIs = GetTickCount
                mRuns(mLastProcedure) = mRuns(mLastProcedure) & ";" & nowIs - mStartProcedureTickTime
                
                If comments <> "" Then
                    For i = UBound(mSession) To 1 Step -1
                        With mSession(i)
                            If .procName = mLastProcedure Then
                                .procComments = .procComments & IIf(.procComments = "", "", "  ||  ") & comments & "<-"
                                Exit For
                            End If
                        End With
                    Next
                End If
                
                incr mActiveCalls, -1
                ReDim Preserve mCallsTree(mActiveCalls)
                
                mLastProcedure = mCallsTree(mActiveCalls).procName
                mStartProcedureTickTime = mCallsTree(mActiveCalls).procStart
            End If
        
        Else
            
            mLastProcedure = IIf(inModule = "", forProcedure, inModule & "." & forProcedure)
            If Not mRuns.exists(mLastProcedure) Then
                Dim runs As String
                runs = ""
                mRuns.add mLastProcedure, runs
            End If
            
            mStartProcedureTickTime = GetTickCount
            
            incr mActiveCalls
            ReDim Preserve mCallsTree(mActiveCalls)
            
            mCallsTree(mActiveCalls).procName = mLastProcedure
            mCallsTree(mActiveCalls).procStart = mStartProcedureTickTime
            
            ReDim Preserve mSession(UBound(mSession) + 1)
            
            With mSession(UBound(mSession))
                .procName = mLastProcedure
                .procComments = IIf(comments = "", "", "->" & comments)
                .procStart = mActiveCalls
            End With
        
        End If
    End If

End Sub

Private Function sysTimeToTime(sTime As SYSTEMTIME) As String
    With sTime
        sysTimeToTime = fromTemplate("{%1%}-{%2%}-{%3%} {%4%}:{%5%}:{%6%}.{%7%}", _
                                    .wYear, Format(.wMonth, "00"), Format(.wDay, "00"), Format(.wHour, "00"), _
                                    Format(.wMinute, "00"), Format(.wSecond, "00"), Format(.wMilliseconds, "000"))
    End With
End Function

Private Function msToTimeString(ByVal duration_ms As Double) As String
    Dim s As Integer, m As Integer, h As Long, last As Long
    
    last = duration_ms
    
    h = last \ 3600000
    last = last - h * 3600000
    
    m = last \ 60000
    last = last - m * 60000
    
    s = last \ 1000
    last = last - s * 1000

    Select Case duration_ms
    Case Is < 1000
        msToTimeString = last & "ms"

    Case Is < 60000
        msToTimeString = fromTemplate("{%1%}s {%2%}ms", s, last)

    Case Is < 3600000
        msToTimeString = fromTemplate("{%1%}min {%2%}s {%3%}ms", m, s, last)

    Case Else
        msToTimeString = fromTemplate("{%1%}h {%2%}min {%3%}s {%4%}ms", h, m, s, last)

    End Select

End Function

Private Sub Class_Initialize()
    Set mRuns = New Dictionary
    Set mLog = New COmniLogger

    mStartProcedureTickTime = 0
    mLastProcedure = ""

    mDurationSession = 0
    mActiveCalls = 0

    mIsSessionStarted = False

    ReDim mSession(0)
    ReDim mCallsTree(0)
    ReDim mTabHeader(0)

End Sub

Private Sub Class_Terminate()
    Set mRuns = New Dictionary
    Set mLog = New COmniLogger
    ReDim mSession(0)
    ReDim mCallsTree(0)
    ReDim mTabHeader(0)

End Sub

Private Sub log(ByVal strToPrint As String)
    mLog.log strToPrint
End Sub


