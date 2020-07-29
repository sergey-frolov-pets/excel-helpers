Attribute VB_Name = "MMultivalues"
Private frm As FMultiValues
Private isLoaded As Boolean

Public Sub ShowMultiValuesBox()
Dim sep As String
Dim old As Variant
Dim o As Variant

If isLoaded Then
    If frm.lstVal.ListCount > 0 Then
        For i = 0 To frm.lstVal.ListCount - 1
           frm.lstVal.Selected(i) = False
        Next
        
        sep = frm.txtSep.text
        old = Split(ActiveCell.Value, sep)
        
        For Each o In old
            For i = 0 To frm.lstVal.ListCount - 1
                If frm.lstVal.List(i) = o Then
                    frm.lstVal.Selected(i) = True
                    Exit For
                End If
            Next
        Next
    End If
Else
    Set frm = New FMultiValues
    isLoaded = True
End If

If Selection.cells.Count > 1 Then
   o = Replace(Selection.Address, "$", "")
   frm.Caption = "Select values for cells {" & o & "}"
Else
   o = Replace(ActiveCell.Address, "$", "")
   
   sep = ActiveCell.Value
   If sep = vbNullString Then sep = "empty"
   
   frm.Caption = "Select values for cell {" & o & "} current value: " & sep
End If

frm.lstVal.SetFocus
frm.Show 1
End Sub
