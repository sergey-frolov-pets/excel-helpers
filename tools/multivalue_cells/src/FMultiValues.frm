VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FMultiValues 
   Caption         =   "Select Values for Cell:"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   OleObjectBlob   =   "FMultiValues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FMultiValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAll_Click()
Dim i As Integer

For i = 0 To lstVal.ListCount - 1
    lstVal.Selected(i) = True
Next
End Sub

Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub btnOK_Click()

Dim selectedValues As String
Dim i As Integer
Dim s As Variant

selectedValues = vbNullString

For i = 0 To lstVal.ListCount - 1
    If lstVal.Selected(i) = True Then selectedValues = selectedValues & lstVal.List(i) & txtSep.text
Next

For Each s In Selection.cells
    
    If selectedValues = vbNullString Then
        s.Value = vbNullString
    Else
        s.Value = Mid(selectedValues, 1, Len(selectedValues) - Len(txtSep.text))
    End If
    
Next

Me.Hide

End Sub

Private Sub btnReset_Click()
Dim i As Integer
        
For i = 0 To lstVal.ListCount - 1
    lstVal.Selected(i) = False
Next

End Sub

Private Sub btnUpdateFromCells_Click()
Dim cell As Variant

On Error Resume Next

lstVal.Clear

For Each cell In Selection.cells
    lstVal.AddItem cell.Value
Next

Me.Hide
End Sub

Private Sub btnUpdateOptions_Click()
Dim varArray As Variant
Dim cell As Variant

On Error Resume Next

varArray = Split(PasteFromClipboard(), vbCrLf)

lstVal.Clear

For Each cell In varArray
    If cell <> vbNullString Then lstVal.AddItem cell
Next

End Sub
