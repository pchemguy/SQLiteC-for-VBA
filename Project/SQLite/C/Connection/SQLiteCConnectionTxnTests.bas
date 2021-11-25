Attribute VB_Name = "SQLiteCConnectionTxnTests"
'@Folder "SQLite.C.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCConnectionTxnTests"
Private TestCounter As Long

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


'''' Txn state should not change between BEGIN DEFERRED and COMMIT.
'''' Txn state should change if IMMEDIATE or EXCLUSIVE are used.
'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnDEFERRED()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_DEFERRED), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnIMMEDIATE()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.IsTrue SQLITE_TXN_NONE < dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnEXCLUSIVE()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_EXCLUSIVE), "Unexpected Txn Begin error"
    Assert.IsTrue SQLITE_TXN_NONE < dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnRead()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim ResultCode As SQLiteResultCodes
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_DEFERRED), "Unexpected Txn Begin error"
    ResultCode = dbc.ExecuteNonQueryPlain("SELECT * FROM functions;")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    Assert.AreEqual SQLITE_TXN_READ, dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginRollback_VerifiesTxnState()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.IsTrue SQLITE_TXN_NONE < dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.Rollback, "Unexpected Txn Rollback error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginRollbackCommit_VerifiesError()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_OK, dbc.Rollback, "Unexpected Txn Rollback error"
    Assert.AreEqual SQLITE_ERROR, dbc.Commit, "Expected SQLITE_ERROR error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommitRollback_VerifiesError()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_ERROR, dbc.Rollback, "Expected SQLITE_ERROR error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcReleasePoint_VerifiesError()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_ERROR, dbc.ReleasePoint("ABCDEFG"), "Expected SQLITE_ERROR error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcSaveRelease_VerifiesTxnState()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.SavePoint("ABCDEFG"), "Unexpected Txn SavePoint error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual SQLITE_OK, dbc.ReleasePoint("ABCDEFG"), "Unexpected Txn ReleasePoint error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcSavepointBeginCommit_VerifiesError()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
Act:
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual SQLITE_OK, dbc.SavePoint("ABCD"), "Unexpected Txn SavePoint error"
    Assert.AreEqual SQLITE_ERROR, dbc.Begin, "Expected SQLITE_ERROR error"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesBusyStatusWithLockingTransaction()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim dbcA As SQLiteCConnection
    Set dbcA = SQLiteCConnection(dbc.DbPathName)
    
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_OK, dbcA.OpenDb, "Unexpected OpenDb error"
    
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_TXN_WRITE, dbc.TxnState("main"), "Unexpected Txn state"
Act:
Assert:
    Assert.AreEqual SQLITE_BUSY, dbcA.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin status"
    Assert.AreEqual SQLITE_TXN_NONE, dbcA.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_OK, dbcA.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesNoTransactionLockingWithMemoryDb()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMemFuncWithData
    Dim dbcA As SQLiteCConnection
    Set dbcA = SQLiteCConnection(dbc.DbPathName)
    
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_OK, dbcA.OpenDb, "Unexpected OpenDb error"
    
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_TXN_WRITE, dbc.TxnState("main"), "Unexpected Txn state"
Act:
Assert:
    Assert.AreEqual SQLITE_OK, dbcA.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin status"
    Assert.AreEqual SQLITE_TXN_WRITE, dbcA.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbcA.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_OK, dbc.Commit, "Unexpected Txn Commit error"
    Assert.AreEqual SQLITE_OK, dbcA.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcSavePointRelease_VerifiesTransactionStates()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    
    Dim ResultCode As SQLiteResultCodes
    Dim SavePointName As String
    Dim AffectedRows As Long
    
    '''' If save point name starts with a digit, the
    '''' "unrecognized token" error is returned.
    SavePointName = "AA" & Left$(GenerateGUID, 8)
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Act:
Assert:
    Assert.AreEqual SQLITE_OK, dbc.SavePoint(SavePointName), "Unexpected SavePoint error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
        
    ResultCode = dbc.ExecuteNonQueryPlain("SELECT * FROM pragma_function_list()")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    Assert.AreEqual SQLITE_TXN_READ, dbc.TxnState("main"), "Unexpected Txn state"
        
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
    Assert.IsTrue AffectedRows > 10, "Unexpected ExecuteNonQueryPlain result"
    Assert.AreEqual SQLITE_TXN_WRITE, dbc.TxnState("main"), "Unexpected Txn state"
    
    ResultCode = dbc.Rollback(SavePointName)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Rollback Txn error"
    If SQLITE_TXN_WRITE <> dbc.TxnState("main") Then Assert.Inconclusive _
        "Previously, Rollback SavePoint failed to reset transaction"
    
    Assert.AreEqual SQLITE_OK, dbc.ReleasePoint(SavePointName), "Unexpected ReleasePoint Txn error"
    Assert.AreEqual SQLITE_TXN_NONE, dbc.TxnState("main"), "Unexpected Txn state"
    
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbc.OpenDb(SQLITE_OPEN_READONLY), "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_DB_READ, dbc.AccessMode, "Unexpected db access mode"
    
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Txn Begin error"
    Assert.AreEqual SQLITE_TXN_READ, dbc.TxnState("main"), "Unexpected Txn state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcDbIsLocked_VerifiesLockState()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim dbcA As SQLiteCConnection
    Set dbcA = SQLiteCConnection(dbc.DbPathName)
    
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
    Assert.AreEqual SQLITE_OK, dbcA.OpenDb, "Unexpected OpenDb error"
Act:
Assert:
    Assert.IsFalse dbc.DbIsLocked, "Unexpected lock state"
    
    Assert.AreEqual SQLITE_OK, dbc.Begin(SQLITE_TXN_IMMEDIATE), "Unexpected Begin Txn error"
    Assert.AreEqual SQLITE_TXN_WRITE, dbc.TxnState("main"), "Unexpected Txn state"
    Assert.AreEqual CVErr(ErrNo.AdoInTransactionErr), dbc.DbIsLocked, "Unexpected lock state"
    Assert.AreEqual True, dbcA.DbIsLocked, "Unexpected lock state"
CleanUp:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
