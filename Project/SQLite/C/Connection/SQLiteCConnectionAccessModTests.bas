Attribute VB_Name = "SQLiteCConnectionAccessModTests"
'@Folder "SQLite.C.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCConnectionAccessModTests"
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
    FixObjC.CleanUp
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesDefaultAccess()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
    ResultCode = dbc.OpenDb(SQLITE_OPEN_DEFAULT)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_FULL, DbAccessMode, "Expected full db access mode"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesReadAccess()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
    ResultCode = dbc.OpenDb(SQLITE_OPEN_READONLY)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_READ, DbAccessMode, "Expected read db access mode"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesDefaultAccessReadOnlyFile()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCReadOnlyAttr
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
    ResultCode = dbc.OpenDb(SQLITE_OPEN_DEFAULT)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_READ, DbAccessMode, "Expected read db access mode"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
