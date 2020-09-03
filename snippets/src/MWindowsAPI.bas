Attribute VB_Name = "MWindowsAPI"
'@Folder("sfSnippets")
''' <summary>
''' --------------------------
''' Module <c>MWindowsAPI</c>
''' --------------------------
''' version 0.1 (2020-08-20)
''' --------------------------
''' Functions from Windows API
''' --------------------------
''' <references>
''' <c>none</c>
''' </references>
''' --------------------------
''' created 2020-08-20
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

''' <summary>
''' --------------------------
''' Sub <c>Sleep</c>
''' --------------------------
''' Make pause in program
''' --------------------------
''' <param><c>dwMilliseconds</c> - Duration in milliseconds</param>
''' --------------------------
''' created 2020-08-26
''' --------------------------
''' </summary>

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

