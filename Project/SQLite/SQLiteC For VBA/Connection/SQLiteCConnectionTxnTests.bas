Attribute VB_Name = "SQLiteCConnectionTxnTests"
'@Folder "SQLite.SQLiteC For VBA.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
Option Explicit
Option Private Module

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
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'''' Txn state should not change between BEGIN DEFERRED and COMMIT.
'''' Txn state should change if IMMEDIATE or EXCLUSIVE are used.
'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnDEFERRED()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_DEFERRED)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnIMMEDIATE()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode > SQLITE_TXN_NONE, "Unexpected Txn state"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnEXCLUSIVE()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_EXCLUSIVE)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode > SQLITE_TXN_NONE, "Unexpected Txn state"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommit_VerifiesTxnRead()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbTempWithFunctionsTableAndData
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_DEFERRED)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    ResultCode = dbc.ExecuteNonQueryPlain("SELECT * FROM functions;")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_READ, TxnStateCode, "Unexpected Txn state"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginRollback_VerifiesTxnState()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode > SQLITE_TXN_NONE, "Unexpected Txn state"
    ResultCode = dbc.Rollback
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Rollback error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.AreEqual SQLITE_TXN_NONE, TxnStateCode, "Unexpected Txn state"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginRollbackCommit_VerifiesError()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    ResultCode = dbc.Rollback
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Rollback error"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcBeginCommitRollback_VerifiesError()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Begin error"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
    ResultCode = dbc.Rollback
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcReleasePoint_VerifiesError()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.ReleasePoint("ABCDEFG")
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcSaveRelease_VerifiesTxnState()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.SavePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn SavePoint error"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode = SQLITE_TXN_NONE, "Unexpected Txn state"
    ResultCode = dbc.ReleasePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn ReleasePoint error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Transactions")
Private Sub ztcSavepointBeginCommit_VerifiesError()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCDbMemory
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.SavePoint("ABCD")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn SavePoint error"
    ResultCode = dbc.Begin
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error"
    ResultCode = dbc.Commit
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn Commit error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
