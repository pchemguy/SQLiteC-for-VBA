Attribute VB_Name = "FixObjCTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "FixObjCTests"
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
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Fixture")
Private Sub ztcGetDBCTmpFuncWithData_VerifiesTmpDatabase()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.FunctionsCount
    Dim Expected As Variant
    Expected = dbs.GetScalar(SQLQuery)
    Assert.IsTrue IsNumeric(Expected), "Unexpected query result."
    SQLQuery = LiteMetaSQL.Create().Tables
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Act:
    SQLQuery = "SELECT count(*) FROM functions"
    Dim Actual As Variant
    Actual = dbs.GetScalar(SQLQuery)
Assert:
    Assert.IsTrue IsNumeric(Actual), "Unexpected query result."
    Assert.AreEqual Expected, Actual, vbNullString
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBCMemFuncWithData_VerifiesMemDatabase()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMemFuncWithData
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Assert.IsTrue dbc.DbHandle > 0, "Expected opened database."
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.FunctionsCount
    Dim Expected As Variant
    Expected = dbs.GetScalar(SQLQuery)
    Assert.IsTrue IsNumeric(Expected), "Unexpected query result."
    
    SQLQuery = LiteMetaSQL.Create().Tables
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
    Assert.IsTrue IsArray(RowSet2D), "Unexpected query result."
    Assert.AreEqual "functions", RowSet2D(0, 0), "Unexpected query result."
Act:
    SQLQuery = "SELECT count(*) FROM functions"
    Dim Actual As Variant
    Actual = dbs.GetScalar(SQLQuery)
Assert:
    Assert.IsTrue IsNumeric(Actual), "Unexpected query result."
    Assert.AreEqual Expected, Actual, vbNullString
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBSMemITRBWithData_VerifiesDBSMemITRBWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbs As ILiteADO
    Set dbs = FixObjC.GetDBSMemITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbs.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBSMemFuncWithData_VerifiesDBSMemFuncWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbs As ILiteADO
    Set dbs = FixObjC.GetDBSMemFuncWithData
    Dim Expected As Long
    Expected = dbs.GetScalar(LiteMetaSQL.FunctionsCount)
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbs.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBSTmpITRBWithData_VerifiesDBSTmpITRBWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbs As ILiteADO
    Set dbs = FixObjC.GetDBSTmpITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbs.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBSTmpFuncWithData_VerifiesDBSTmpFuncWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbs As ILiteADO
    Set dbs = FixObjC.GetDBSTmpFuncWithData
    Dim Expected As Long
    Expected = dbs.GetScalar(LiteMetaSQL.FunctionsCount)
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbs.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
