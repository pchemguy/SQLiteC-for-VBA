Attribute VB_Name = "SQLiteCStatementTests"
'@Folder "SQLiteC For VBA.Statement"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If
Private FixObj As SQLiteCTestFixObj
Private FixSQL As SQLiteCTestFixSQL


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set FixObj = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DbStatement")
Private Sub ztcCreateStatement_VerifiesNewStatement()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Assert:
    Assert.IsNotNothing dbs, "DbStmt is not set."
    Assert.IsNotNothing dbs.DbConnection, "Connection object not set."
    Assert.IsNotNothing dbs.DbExecutor, "Executor object not set."
    Assert.IsNothing dbs.DbParameters, "Parameters object should not be set."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    Assert.AreSame dbc, dbs.DbConnection, "Connection object mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_ThrowsOnClosedConnection()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTSQLiteVersion
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesPrepareSQLiteVersion()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTSQLiteVersion
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, dbs.StmtHandle, "StmtHandle should not be zero."
    Assert.IsNotNothing dbs.DbParameters, "Parameters object should be set."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, dbs.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesPrepareOfCreateTable()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.CREATETableITRB
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, dbs.StmtHandle, "StmtHandle should not be zero."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, dbs.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesErrorOnInvalidSQL()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    
    SQLQuery = "SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "-- SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Expected SQLITE_OK result: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "ABC SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "SELECT * FROM ABC"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesErrorWithSelectFromFakeTable()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemoryWithTable
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    
    SQLQuery = FixSQL.SELECTTestTable
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcGetBusy_VerifiesBusyStatus()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTCollations
    
        ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual False, dbs.Busy, "Busy status should be False"
        Result = dbs.GetScalar(SQLQuery)
    Assert.AreEqual True, dbs.Busy, "Busy status should be True"
        ResultCode = dbs.Reset
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Reset error."
    Assert.AreEqual False, dbs.Busy, "Busy status should be False"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesGetScalar()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.GetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTSQLiteVersion
    
    Result = dbs.GetScalar(SQLQuery)
    Assert.AreEqual dbm.Version(False), Result, "GetScalar mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnBlankQueryAndNullParams()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQueryDummy As String
    SQLQueryDummy = FixSQL.SELECTSQLiteVersion
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = Null
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQueryDummy)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnInvalidParamsType()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQueryDummy As String
    SQLQueryDummy = FixSQL.SELECTSQLiteVersion
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = "ABC"
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQueryDummy)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnBlankQueryToUnpreparedStatement()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = Array("ABC")
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query Paged RowSet")
Private Sub ztcGetPagedRowSet_VerifyPageRowSetGeometry()
    On Error GoTo TestFail

    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim PageSize As Long
    PageSize = 8
    Dim PageCount As Long
    PageCount = 28
    dbs.DbExecutor.PageSize = PageSize
    dbs.DbExecutor.PageCount = PageCount
    Dim AffectedRows As Long
    AffectedRows = FixObj.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTMinMaxSubstrLTrimFromFunctionsNamedParam
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = FixSQL.SELECTMinMaxSubstrLTrimFunctionsNamedValues
    Dim PagedRowSet As Variant
    PagedRowSet = dbs.GetPagedRowSet(SQLQuery, QueryParams, True)
Assert:
    Assert.IsFalse IsError(PagedRowSet), "Unexpected error from GetPagedRowSet."
    Assert.IsFalse IsEmpty(PagedRowSet), "GetPagedRowSet should not be empty."
    Assert.IsFalse IsNull(PagedRowSet), "GetPagedRowSet should not be null."
    Assert.AreEqual 0, LBound(PagedRowSet), "PagesArray base mismatch"
    Assert.AreEqual PageCount - 1, UBound(PagedRowSet), "PagesArray size mismatch"
    Assert.AreEqual 0, LBound(PagedRowSet(0)), "RowSet base mismatch"
    Assert.AreEqual PageSize - 1, UBound(PagedRowSet(0)), "RowSet size mismatch"
    Assert.AreEqual 0, LBound(PagedRowSet(0)(0)), "FieldSet base mismatch"
    Assert.AreEqual dbs.DbExecutor.GetColumnCount - 1, UBound(PagedRowSet(0)(0)), "FieldSet size mismatch"
Cleanup:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query Paged RowSet")
Private Sub ztcGetPagedRowSet_SelectSubsetOfFunctions()
    On Error GoTo TestFail

    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    'Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    dbs.DbExecutor.PageSize = 99
    dbs.DbExecutor.PageCount = 9
    Dim AffectedRows As Long
    AffectedRows = FixObj.CreateFunctionsTableWithData(dbc)

    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTMinMaxSubstrLTrimFromFunctions
    Dim SQLQueryCount As String
    SQLQueryCount = FixSQL.CountSelectNoCTE(SQLQuery)
    Dim RecordCount As Variant
    RecordCount = dbs.GetScalar(SQLQueryCount)
Act:
    SQLQuery = FixSQL.SELECTMinMaxSubstrLTrimFromFunctionsNamedParam
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = FixSQL.SELECTMinMaxSubstrLTrimFunctionsNamedValues
    Dim PagedRowSet As Variant
    PagedRowSet = dbs.GetPagedRowSet(SQLQuery, QueryParams, True)
Assert:
    Assert.IsFalse IsEmpty(PagedRowSet(0)(RecordCount - 1)), "RowSet is too small"
    Assert.IsTrue IsEmpty(PagedRowSet(0)(RecordCount)), "RowSet is too big"
Cleanup:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

