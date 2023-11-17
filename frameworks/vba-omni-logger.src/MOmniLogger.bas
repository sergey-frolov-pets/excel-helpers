Attribute VB_Name = "MOmniLogger"
'@Folder("OmniLogger framework")
''' <summary>
''' --------------------------
''' Module <c>MOmniLogger.bas</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' OPTIONAL OmniLogger framework file to make logging code more compact
''' --------------------------
''' <references>
''' <c>COmniLogger.cls</c>
''' <c>MSupport.bas</c>
''' </references>
''' --------------------------
''' created 2023-09-01
''' by Sergey Frolov (pet-projects@sergey-frolov.ru)
''' --------------------------
''' </summary>
'''
''' <license>
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see
''' https://www.gnu.org/licenses/
''' </license>

Option Explicit

Private mLogger As COmniLogger

''' <summary>
''' --------------------------
''' Property <c>logger</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' logger is COmniLogger object initiated by initLogger() sub.
''' Automatically will create logDebug type stream if will not be initiated before the first call.
''' --------------------------
''' </summary>
Public Property Get logger() As COmniLogger
    'by default logging stream is Immediate window: option = logDebug
    If mLogger Is Nothing Then mLogger.initLog logDebug

    Set logger = mLogger

End Property

''' <summary>
''' --------------------------
''' Sub <c>initLogger</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Wrapper for COmniLogger.initLog() sub (see description there)
''' --------------------------
''' </summary>
Public Sub initLogger(ByVal targetType As LOG_STREAM_TYPE, Optional target As String, Optional newHeader As String = "", Optional delimiter As String = vbTab, Optional withTimeStamp As Boolean = False, Optional timestampFormat As String = "yyyy-mm-dd hh:mm:ss")
    Set mLogger = Nothing
    Set mLogger = New COmniLogger
    
    mLogger.initLog targetType, target, newHeader, delimiter, withTimeStamp, timestampFormat
End Sub

''' <summary>
''' --------------------------
''' Sub <c>log</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Wrapper for COmniLogger.log() sub (see description there)
''' --------------------------
''' </summary>
Public Sub log(ParamArray data() As Variant)
    If UBound(data) = 0 Then
        mLogger.log data(0)
    Else
        Dim arr As Variant, i As Integer, d As Variant
        ReDim arr(UBound(data))
        i = 0
        For Each d In data
            arr(i) = d
            i = i + 1
        Next
        mLogger.logArray arr
        ReDim arr(0)
    End If
End Sub

''' <summary>
''' --------------------------
''' Sub <c>logArray</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Wrapper for COmniLogger.logArray() sub (see description there)
''' --------------------------
''' </summary>
Public Sub logArray(data As Variant)
    mLogger.logArray data
End Sub


''' <summary>
''' --------------------------
''' Sub <c>logHeader</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Wrapper for COmniLogger.logHeader() sub (see description there)
''' --------------------------
''' </summary>
Public Sub logHeader(ParamArray data() As Variant)
    mLogger.logHeader
End Sub
