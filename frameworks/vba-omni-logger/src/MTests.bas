Attribute VB_Name = "MTests"
'@Folder("tests")

''' <summary>
''' --------------------------
''' Module <c>MTests.bas</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Module with examples how to use OmniLogger framework
''' DO NOT COPY it to your project!
''' Run Sub testOmniLogger() to see how OmniLogger works
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

'logging data for examples
Dim mHeader As String
Dim mArr As Variant
Dim mInfo As String

'logging data preparation
Private Sub init_test_data()
    mHeader = "[Col 1]" & vbTab & "[Col 2]" & vbTab & "[Col 3]" & vbTab & "[1]" & vbTab & "[2]" & vbTab & "[3]"
    mArr = Array("Val A", "Val B", "Val C", 1, 2, 3)
    mInfo = "Val D" & vbTab & "Val E" & vbTab & "Val F" & vbTab & 4 & vbTab & 5 & vbTab & 6

End Sub

''' <summary>
''' --------------------------
''' Sub <c>testOmniLogger</c>
''' --------------------------
''' version 1.0 (2023-09-01)
''' --------------------------
''' Run this sub to see how OmniLogger works
''' --------------------------
''' </summary>

Public Sub testOmniLogger()
   
'preparing data for test
    init_test_data

'[ATTENTION!] Choose only one log stream for testing
    initDebugLog 'check Immediate Window for results
    'initRangeLog 'check [Log] sheet for results
    'initTXTLog 'check "C:\!delMeNow_OmniLoggerExample.txt" file for results
    
    If logger.withTimeStamp Then logger.header = "[Timestamp]" & vbTab & mHeader

    logger.clear
    logger.delimiter = vbTab
    
    logHeader
    
    logArray mArr
    log mInfo
    log "Val G", "Val H", "Val I", 7, 8, 9
    
    logger.closeLog
    
    'one more logger
    Dim l2 As New COmniLogger
    l2.initLog logDebug, , , , True, "hh:mm"
    l2.log "Logging into parallel stream (l2)"
    Set l2 = Nothing
    
    MsgBox "Check [" & logger.logTarget & "] for results.", vbInformation, "Done!"
End Sub

'init logger to log into Immediate window
Private Sub initDebugLog()
    initLogger logDebug, , mHeader, vbTab, True, "hh:nn:ss"
    
    log "Testing clear(), log(), logHeader() and logArray()"
    log "--------------------------------------------------"
    logger.clear
    
    logger.delimiter = " "
    log "sub clear()", "adds two empty lines", "to", "debug window"

End Sub

'init logger to log into selected Range
Private Sub initRangeLog()
    initLogger logRange, "Log!B2", mHeader, , True

End Sub

'init logger to log into txt/csv file
Private Sub initTXTLog()
    '!!! options logCSV and logTXT are synonyms
    initLogger logCSV, fld & "C:\!delMeNow_OmniLoggerExample.txt", mHeader
    
    logger.openLog
    log "something as file header"

End Sub
