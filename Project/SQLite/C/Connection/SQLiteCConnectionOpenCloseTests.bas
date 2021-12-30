Attribute VB_Name = "SQLiteCConnectionOpenCloseTests"
'@Folder "SQLite.C.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCConnectionOpenCloseTests"
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


'@TestMethod("Connection")
Private Sub ztcCreateConnection_VerifiesSQLiteCConnectionWithValidDbPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCReg
Assert:
    Assert.IsFalse dbc Is Nothing, "Default SQLiteCConnection is not set."
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbPathName_VerifiesMemoryDbPathName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Assert:
    Assert.AreEqual DbPathName, dbc.DbPathName
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbPathName_VerifiesAnonDbPathName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim DbPathName As String
    DbPathName = vbNullString
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCAnon
Assert:
    Assert.AreEqual DbPathName, dbc.DbPathName
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcAttachedDbPathName_ThrowsOnClosedConnection()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Debug.Print dbc.DbPathName = dbc.AttachedDbPathName
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("Connection")
Private Sub ztcAttachedDbPathName_VerifiesTempDbPathName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual dbc.DbPathName, dbc.AttachedDbPathName, "AttachedDbPathName mismatch."
CleanUp:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithRegularDb()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCReg
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithTempDb()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCAnon
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithMemoryDb()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesErrInfo()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCReg
    Dim ErrInfo As SQLiteCErr
    ErrInfo = dbc.ErrInfo
Act:
Assert:
    With ErrInfo
        Assert.AreEqual SQLITE_OK, .ErrorCode, "ErrorCode mismatch"
        Assert.AreEqual SQLITE_OK, .ErrorCodeEx, "ErrorCodeEx mismatch"
        Assert.AreEqual vbNullString, .ErrorName, "ErrorName mismatch"
        Assert.AreEqual vbNullString, .ErrorCodeName, "ErrorCodeName mismatch"
        Assert.AreEqual vbNullString, .ErrorCodeExName, "ErrorCodeExName mismatch"
        Assert.AreEqual vbNullString, .ErrorMessage, "ErrorMessage mismatch"
        Assert.AreEqual vbNullString, .ErrorString, "ErrorString mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcGetErr_ThrowsOnClosedConnection()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCReg
    dbc.ErrInfoRetrieve
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub
