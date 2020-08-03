VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcelSpamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iSubj As Range
Private iBody As Range
Private iSender As Range
Private iParams As Range
Private iAttachmentPath As Range
Private iImportance As Range
Private paramName() As String
Private iOutApp As Object
    
Public Sub initSpamCells(subjectCell As Range, bodyCell As Range, paramFirstCells As Range, importanceCell As Range, senderCell As Range, attachmentPathCell As Range)
Set iSubj = subjectCell
Set iBody = bodyCell
Set iSender = senderCell
Set iAttachmentPath = attachmentPathCell
Set iImportance = importanceCell
Set iParams = paramFirstCells
End Sub

Sub sendSpam()
    Dim i As Integer
    Dim recipients As Range
    Dim recipient As Range
    
enableFastCode True
    
    Range(iParams.Cells(1, 1), iParams.Cells(1, 1).SpecialCells(xlLastCell)).Select
    Set recipients = Selection
    
    ReDim paramName(recipients.Columns.Count)
    
    For i = 1 To recipients.Columns.Count
        If recipients.Cells(1, i).Value = vbNullString Then Exit For
        paramName(i) = recipients.Cells(1, i).Value
    Next
    ReDim Preserve paramName(i - 1)
    
    Set iOutApp = CreateObject("Outlook.Application")
    
    For i = 2 To recipients.Rows.Count
        Set recipient = recipients.Rows(i)
        If recipient.Cells(1, 1).Value = vbNullString Then Exit For
        sendEmail recipient
    Next

    Set iOutApp = Nothing
    
    enableFastCode False
    
    MsgBox i - 2 & " emails were sent.", vbInformation, "Done"
End Sub

Private Sub sendEmail(recipient As Range)
    Dim email As String, subj As String
    Dim body As String, url As String
    Dim path2Attachments As String, attach As String
    Dim i As Integer, r As Integer, x As Double
    Dim importance As Integer
    Dim cc As String, bcc As String
    
    body = iBody.Cells(1, 1).Value
    subj = iSubj.Cells(1, 1).Value
    email = recipient.Cells(1, 1).Value
    cc = vbNullString
    bcc = vbNullString
    attach = vbNullString
    
    For i = 2 To UBound(paramName)
        If LCase(paramName(i)) = "cc" Then cc = recipient.Cells(1, i).Value
        If LCase(paramName(i)) = "bcc" Then bcc = recipient.Cells(1, i).Value
        If LCase(paramName(i)) = "attachment" Or LCase(paramName(i)) = "приложение" Then attach = recipient.Cells(1, i).Value
        subj = Replace(subj, "{" & paramName(i) & "}", recipient.Cells(1, i).Value)
        body = Replace(body, "{" & paramName(i) & "}", recipient.Cells(1, i).Value)
    Next
        
    On Error Resume Next
    
    Dim outMail As Object
    Set outMail = iOutApp.CreateItem(0)
    
    If iAttachmentPath.Value <> vbNullString Then
        If Mid(iAttachmentPath.Value, Len(iAttachmentPath.Value), 1) <> "\" Then
            path2Attachments = iAttachmentPath.Value & "\"
        Else
            path2Attachments = iAttachmentPath.Value
        End If
    End If
    
    With outMail
        If iSender.Value <> "" Then .SentOnBehalfOfName = iSender.Value
        .To = email
        .cc = cc
        .bcc = bcc
        .subject = subj
        .body = body
        .importance = Val(iImportance.Cells(1, 1).Value)
        If path2Attachments <> vbNullString Then
            If attach <> vbNullString Then .Attachments.Add path2Attachments & attach
        End If
        .Send
     End With
    
    Set outMail = Nothing
End Sub
