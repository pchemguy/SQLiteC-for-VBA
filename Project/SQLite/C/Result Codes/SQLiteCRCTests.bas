Attribute VB_Name = "SQLiteCRCTests"
'@Folder "SQLite.C.Result Codes"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCCRCTests"
Private TestCounter As Long

#Const LateBind = 0     '''' RubberDuck Tests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("ResultCodes")
Private Sub ztcCodeToName_VerifiesCodeName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
Assert:
    Assert.AreEqual "OK", SQLiteCRC.CodeToName(SQLITE_OK), "OK name mismatch"
    Assert.AreEqual "ERROR", SQLiteCRC.CodeToName(SQLITE_ERROR), "ERROR name mismatch"
    Assert.AreEqual "IOERR_BEGIN_ATOMIC", SQLiteCRC.CodeToName(SQLITE_IOERR_BEGIN_ATOMIC), "IOERR_BEGIN_ATOMIC name mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
