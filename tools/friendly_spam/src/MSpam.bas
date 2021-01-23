Attribute VB_Name = "MSpam"
Option Explicit


Sub btnSendSpamWithDelay()
If [B1] <> vbNullString Then
    If MsgBox([B1] & vbCrLf & vbCrLf & "Repeat?", vbYesNo, "Repeat sending") = vbNo Then Exit Sub
End If

Dim ss As New CExcelSpamer
Dim on_time As Date

on_time = 0

If [D2] > 0 Then
    Select Case Trim(LCase([E2]))
    Case "minutes":
        If [D2] > 1440 * 7 Then GoTo fin
        on_time = DateAdd("n", [D2], Now())
        
    Case "hours":
        If [D2] > 24 * 14 Then GoTo fin
        on_time = DateAdd("h", [D2], Now())
    
    Case "days":
        If [D2] > 30 Then GoTo fin
        on_time = DateAdd("d", [D2], Now())
    
    Case Else:
        GoTo fin
    End Select
        
    If on_time = 0 Then GoTo fin
    
    ss.initSpamCells [B7], [B10], [B22], [C4], [G2], [F4]
    ss.sendSpam on_time
    
    [B1] = "Email will be sent to the mailing list at " & Format(on_time, "hh:mm (dd.mm.yyyy)")
    Exit Sub
End If

fin:
    MsgBox "Check delay parameters - current [" & [D2] & " " & [E2] & "] is not valid. May be too big delay (for 'minutes' + 7 days, for 'hours' - 14 days, for 'days' - 30 days)", vbExclamation, "Emails wasn't send"
End Sub

Sub btnSendSpam()
If [B1] <> vbNullString Then
    If MsgBox([B1] & vbCrLf & vbCrLf & "Repeat?", vbYesNo, "Repeat sending") = vbNo Then Exit Sub
End If

Dim ss As New CExcelSpamer
ss.initSpamCells [B7], [B10], [B22], [C4], [G2], [F4]
ss.sendSpam

[B1] = "Email was sent to the mailing list at " & Format(Now(), "dd.mm.yyyy hh:mm")
End Sub

Sub btnUpdateStatus()

Dim files() As String
Dim i As Integer, l As Integer
Dim testStr As String
Dim attachmentColumn As String

enableFastCode True

[D2] = ""
deleteAllRowsBelowCell [A23]
copyPasteRange Sheets(1).Rows("22:" & lastRowInColumn("A", 1)), [A22]

attachmentColumn = getAttachmentColumn()

If attachmentColumn = vbNullString Then
   MsgBox "There are no attachments for this email!", vbExclamation, "Files were not sent to recipients"
   Exit Sub
End If

files = getFilesList(ActiveWorkbook.Path & [F5], [K5], False)

For i = 1 To UBound(files)
    incr testStr, LCase(Dir(files(i))) & ";"
Next

l = lastRowInColumn("A")
For i = l To 23 Step -1
    If InStr(testStr, LCase(Range(attachmentColumn & i).Value) & ";") > 0 Then Rows(i).Delete
Next

[C4].Activate

i = l - 22
l = lastRowInColumn("A") - 22

enableFastCode False

MsgBox "Replays " & i - l & ", not answered " & l & " out of " & i & " requests.", vbInformation, "Status"

End Sub

Public Function getAttachmentColumn() As String
Dim cell

For Each cell In Range([B22], Rows(22).SpecialCells(xlLastCell))
    If LCase(cell) = "attachment" Or LCase(cell) = "приложение" Then
        getAttachmentColumn = columnNameByIndex(cell.Column)
        Exit Function
    End If
Next
getAttachmentColumn = vbNullString

End Function
