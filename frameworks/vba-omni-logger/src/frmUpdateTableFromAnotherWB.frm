VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateTableFromAnotherWB 
   Caption         =   "Update table records from another workbook"
   ClientHeight    =   9260
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   12560
   OleObjectBlob   =   "frmUpdateTableFromAnotherWB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateTableFromAnotherWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub saveFormState()
Dim i As Integer

ReDim formState(Me.Controls.Count)

For i = 0 To Me.Controls.Count - 1
    If TypeName(Me.Controls(i)) = "TextBox" Then
        formState(i) = Me.Controls(i).text
    
    ElseIf TypeName(Me.Controls(i)) = "ListBox" Then
        formState(i) = Me.Controls(i).ListIndex
    
    ElseIf TypeName(Me.Controls(i)) = "CheckBox" Then
        formState(i) = Me.Controls(i).Value
    
    Else
    
    End If

Next

End Sub

Public Sub loadFormState()

If (Not Not formState) = 0 Then Exit Sub

Dim i As Integer

For i = 0 To Me.Controls.Count - 1
    If TypeName(Me.Controls(i)) = "TextBox" Then
        Me.Controls(i).text = formState(i)
    
    ElseIf TypeName(Me.Controls(i)) = "ListBox" Then
        If Me.Controls(i).ListCount > formState(i) Then Me.Controls(i).ListIndex = formState(i)
    
    ElseIf TypeName(Me.Controls(i)) = "CheckBox" Then
        Me.Controls(i).Value = formState(i)
    
    Else

    End If
Next


End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnPreview_Click()
'update_table_from_another_workbook lstSrcWB.List(lstSrcWB.ListIndex), lstSsht.List(lstSsht.ListIndex), txtSIDs, txtSData, lstDestWB.List(lstDestWB.ListIndex), lstDsht.List(lstDsht.ListIndex), txtDIDs, txtDData, Val(txtSStart), Val(txtSEnd), chkSVisible, chkDVisible, chkDEmpty

preview_values lstSPrev, lstSrcWB.List(lstSrcWB.ListIndex), lstSsht.List(lstSsht.ListIndex), txtSIDs, txtSData, txtSStart, txtSStart + 5
preview_values lstDPrev, lstDestWB.List(lstDestWB.ListIndex), lstDsht.List(lstDsht.ListIndex), txtDIDs, txtDData, txtDStart, txtDStart + 5

End Sub

Private Sub btnUpdate_Click()
If MsgBox("Start updating cells in sheet " & vbCrLf & txtDestWB & vbCrLf & " by values from sheet " & vbCrLf & txtSrcWB & "?", vbQuestion + vbYesNo, "Please, confirm operation") = vbNo Then Exit Sub
update_table_from_another_workbook lstSrcWB.List(lstSrcWB.ListIndex), lstSsht.List(lstSsht.ListIndex), txtSIDs, txtSData, lstDestWB.List(lstDestWB.ListIndex), lstDsht.List(lstDsht.ListIndex), txtDIDs, txtDData, Val(txtSStart), Val(txtSEnd), chkSVisible, chkDVisible, chkDEmpty

End Sub

Private Sub lstDestWB_Click()
lstDsht.Clear
For Each s In Workbooks(lstDestWB.List(lstDestWB.ListIndex)).Sheets
    lstDsht.AddItem s.Name
Next
lstDsht.ListIndex = 0
lstDsht_Click
End Sub

Private Sub lstDPrev_Click()
txtExpand = lstDPrev.List(lstDPrev.ListIndex)

End Sub

Private Sub lstDsht_Click()
txtDestWB = "'[" & lstDestWB.List(lstDestWB.ListIndex) & "]" & lstDsht.List(lstDsht.ListIndex) & "'!"
End Sub

Private Sub lstSPrev_Click()
txtExpand = lstSPrev.List(lstSPrev.ListIndex)
End Sub

Private Sub lstSrcWB_Click()
lstSsht.Clear
For Each s In Workbooks(lstSrcWB.List(lstSrcWB.ListIndex)).Sheets
    lstSsht.AddItem s.Name
Next
lstSsht.ListIndex = 0
lstSsht_Click
End Sub

Private Sub lstSsht_Click()
txtSrcWB = "'[" & lstSrcWB.List(lstSrcWB.ListIndex) & "]" & lstSsht.List(lstSsht.ListIndex) & "'!"
End Sub

Private Sub UserForm_Terminate()
saveFormState
End Sub
