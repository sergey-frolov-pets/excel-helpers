Attribute VB_Name = "MProfiler"
'@Folder("sfProfiler")

Option Explicit

Private mP As CProfiler

Public Property Get p() As CProfiler

If mP Is Nothing Then Set mP = New CProfiler
Set p = mP

End Property


Public Sub p_(Optional forProcedure As String, Optional inModule As String, Optional comments As String)
p.p_ forProcedure, inModule, comments
End Sub
