Attribute VB_Name = "LiteACIDTests"
'@Folder "SQLite.Checks"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, AssignmentNotUsed, VariableNotUsed
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LiteACIDTests"
Private TestCounter As Long

#Const LateBind = 1     '''' RubberDuck Tests
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
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxIntegrityADODB(Optional ByVal FilePathName As String = vbNullString) As Boolean
    Dim DbPathName As String
    DbPathName = IIf(Len(FilePathName) > 0, FilePathName, FixObjAdo.DefaultDbPathName)
    zfxIntegrityADODB = LiteACID(LiteMan(DbPathName).ExecADO).IntegrityADODB
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_PassesDefaultDatabaseIntegrityCheck()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
Assert:
    Assert.IsTrue zfxIntegrityADODB(), "Integrity check on default database failed"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnFileNotDatabase()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Dim CheckResult As Boolean
    CheckResult = zfxIntegrityADODB(ThisWorkbook.Name)
    Guard.AssertExpectedError Assert, ErrNo.OLE_DB_ODBC_Err
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnCorruptedDatabase()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Dim CheckResult As Boolean
    CheckResult = zfxIntegrityADODB(FixObjAdo.RelPrefix & "ICfailFKCfail.db")
    Guard.AssertExpectedError Assert, ErrNo.IntegrityCheckErr
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnFailedFKCheck()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Dim CheckResult As Boolean
    CheckResult = zfxIntegrityADODB(FixObjAdo.RelPrefix & "ICokFKCfail.db")
    Guard.AssertExpectedError Assert, ErrNo.ConsistencyCheckErr
End Sub
