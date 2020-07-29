Attribute VB_Name = "MClipboard"
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Private Const RAISE_ERROR = False

Public Function CopyToClipboard(myString As String) As Boolean
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, hTmp As Long

   ' Allocate moveable global memory.
   hGlobalMemory = GlobalAlloc(GHND, Len(myString) + 1)

   ' Lock the block to get a far pointer to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, myString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      If RAISE_ERROR Then
        Err.Raise 520, "CopyToClipboard", "Could not unlock memory location."
      Else
        CopyToClipboard = False
        GoTo mrk
      End If
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      If RAISE_ERROR Then
        Err.Raise 521, "CopyToClipboard", "Could not open the Clipboard."
      Else
        CopyToClipboard = False
        Exit Function
      End If
   End If

   ' Clear the Clipboard.
   hTmp = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

mrk:

   If CloseClipboard() = 0 Then
        If RAISE_ERROR Then
           Err.Raise 521, "CopyToClipboard", "Could not close the Clipboard."
        Else
            CopyToClipboard = False
            Exit Function
        End If
   Else
        CopyToClipboard = True
   End If

   End Function

Public Function PasteFromClipboard() As String
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim myString As String
   Dim RetVal As Long

   If OpenClipboard(0&) = 0 Then
      If RAISE_ERROR Then
        Err.Raise 521, "PasteFromClipboard", "Could not open the Clipboard."
      Else
        PasteFromClipboard = vbNullString
        Exit Function
      End If
   End If

   ' Obtain the handle to the global memory block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   
   If IsNull(hClipMemory) Then
       If RAISE_ERROR Then
         Err.Raise 520, "PasteFromClipboard", "Could not allocate memory."
       Else
         PasteFromClipboard = False
         GoTo mrk
       End If
   End If

   ' Lock Clipboard memory so we can reference the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      myString = Space$(MAXSIZE)
      RetVal = lstrcpy(myString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)

      ' Peel off the null terminating character.
      myString = Mid(myString, 1, InStr(1, myString, Chr$(0), 0) - 1)
   Else
       If RAISE_ERROR Then
         Err.Raise 520, "PasteFromClipboard", "Could not lock memory to copy string from."
       Else
         PasteFromClipboard = False
         GoTo mrk
       End If
   End If

mrk:

   RetVal = CloseClipboard()
   PasteFromClipboard = myString

End Function
