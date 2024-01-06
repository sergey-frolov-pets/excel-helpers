Attribute VB_Name = "MTests"
'@Folder("Tests")

Option Explicit

Public Sub testProfilerAndReportTypes()
    Dim i As Integer
        
    p.startSession
    
    p_ "testProfilerAndReportTypes", "MTestCProfiler"

    testProfilerProc1
    
    For i = 1 To 2
        testProfilerProc2
    Next
    
    testProfilerProc1
    
    p_
    
    With p
        .stopSession
        
        .toDebug
        .toDebug True
        .toTXT ActiveWorkbook.Path & "\logs\" & "report.txt" '!!! Create 'logs' directory to avoid errors
        .toJSON ActiveWorkbook.Path & "\logs\" & "sessions.json", True
        .toCSV ActiveWorkbook.Path & "\logs\" & "sessions.csv", True
        .toRange "Log!A1", True

    End With

End Sub

Private Sub testProfilerProc1()
    p_ "testProfilerProc1"
    Dim j As Double, i As Long
    For i = 0 To 1000000
        j = i * Rnd
    Next
    p_ , , "p1 done!"
End Sub

Private Sub testProfilerProc2()
    p_ "testProfilerProc2"
        Dim j As Double
        testProfilerProc3 5
        testProfilerProc3 10
    p_ , , "p2 closed"
End Sub

Private Sub testProfilerProc3(ByVal f As Integer)
    p_ "testProfilerProc3", , "f =" & f
    Dim j As Double, i As Long
    For i = 0 To 10000000
        j = f * i * Rnd
    Next
    p_
End Sub


