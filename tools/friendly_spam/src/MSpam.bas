Attribute VB_Name = "MSpam"
Option Explicit

Sub btnSendSpam()
If [D2] <> vbNullString Then
    If MsgBox([D2] & vbCrLf & vbCrLf & "Repeat?", vbYesNo, "Repeat sending") = vbNo Then Exit Sub
End If

Dim ss As New CExcelSpamer
ss.initSpamCells [B7], [B10], [B22], [C4], [G2], [F4]
ss.sendSpam

[D2] = "Email was sent to the mailing list at " & Format(Now(), "dd.mm.yyyy hh:mm")
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
