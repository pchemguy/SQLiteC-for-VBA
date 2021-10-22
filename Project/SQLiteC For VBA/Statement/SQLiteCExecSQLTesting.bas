Attribute VB_Name = "SQLiteCExecSQLTesting"
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


'@TestMethod("Data types")
Private Sub ztcSQLiteTypeName_VerifiesSQLiteTypeName()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.GetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
Assert:
    Assert.AreEqual "INTEGER", dbs.DbExecutor.SQLiteTypeName(SQLITE_INTEGER), "SQLiteTypeName mismatch."
    Assert.AreEqual "FLOAT", dbs.DbExecutor.SQLiteTypeName(SQLITE_FLOAT), "SQLiteTypeName mismatch."
    Assert.AreEqual "TEXT", dbs.DbExecutor.SQLiteTypeName(SQLITE_TEXT), "SQLiteTypeName mismatch."
    Assert.AreEqual "NULL", dbs.DbExecutor.SQLiteTypeName(SQLITE_NULL), "SQLiteTypeName mismatch."
    Assert.AreEqual "BLOB", dbs.DbExecutor.SQLiteTypeName(SQLITE_BLOB), "SQLiteTypeName mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcSQLiteTypeAffinityName_VerifiesSQLiteTypeAffinityName()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.GetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
Assert:
    Assert.AreEqual "BLOB", dbs.DbExecutor.SQLiteTypeAffinityName(SQLITE_AFF_BLOB), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "INTEGER", dbs.DbExecutor.SQLiteTypeAffinityName(SQLITE_AFF_INTEGER), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "NUMERIC", dbs.DbExecutor.SQLiteTypeAffinityName(SQLITE_AFF_NUMERIC), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "REAL", dbs.DbExecutor.SQLiteTypeAffinityName(SQLITE_AFF_REAL), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "TEXT", dbs.DbExecutor.SQLiteTypeAffinityName(SQLITE_AFF_TEXT), "SQLiteTypeAffinityName mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcTypeAffinityFromDeclaredType_VerifiesDeclaredTypeHandling()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.GetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
Assert:
    Assert.AreEqual SQLITE_AFF_INTEGER, dbs.DbExecutor.TypeAffinityFromDeclaredType("UNSIGNED BIG iNt"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_TEXT, dbs.DbExecutor.TypeAffinityFromDeclaredType("NATIVE cHaRACTER(70)"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_BLOB, dbs.DbExecutor.TypeAffinityFromDeclaredType("BLoB"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_REAL, dbs.DbExecutor.TypeAffinityFromDeclaredType("DOuBLE PRECISION"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_NUMERIC, dbs.DbExecutor.TypeAffinityFromDeclaredType("STRING"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_INTEGER, dbs.DbExecutor.TypeAffinityFromDeclaredType("FLoATING POInT"), "TypeAffinityFromDeclaredType mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcTypeAffinityMap_VerifiesMappingToSQLiteTypes()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.GetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
Assert:
    Assert.AreEqual SQLITE_BLOB, dbs.DbExecutor.AffinityMap(SQLITE_AFF_BLOB - &H41), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_TEXT, dbs.DbExecutor.AffinityMap(SQLITE_AFF_TEXT - &H41), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_TEXT, dbs.DbExecutor.AffinityMap(SQLITE_AFF_NUMERIC - &H41), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_INTEGER, dbs.DbExecutor.AffinityMap(SQLITE_AFF_INTEGER - &H41), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_FLOAT, dbs.DbExecutor.AffinityMap(SQLITE_AFF_REAL - &H41), "AffinityMap mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Metadata")
Private Sub ztcGetColumnMetaAPI_VerifiesFunctionsColumnMeta()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsPragmaTable
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
Act:
    '''' Enable this call to obtain a meaningful value for .DataType
    'ResultCode = dbs.DbExecutor.ExecuteStepAPI
    'Assert.AreEqual SQLITE_ROW, ResultCode, "Unexpected ExecuteStepAPI error."
    Dim ColumnInfo As SQLiteCColumnMeta
    ColumnInfo.ColumnIndex = 0
    ColumnInfo.Initialized = -1
    '''' table_column_metadata API against SELECT-PRAGMA should fail.
    ResultCode = dbs.DbExecutor.ColumnMetaAPI(ColumnInfo)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Unexpected GetColumnMetaAPI error."
Assert:
    With ColumnInfo
        Assert.AreEqual SQLITE_AFF_INTEGER, .Affinity, "Affinity mismatch."
        Assert.AreEqual SQLITE_NULL, .DataType, "DataType mismatch"
        Assert.AreEqual "main", .DbName, "Db alias mismatch."
        Assert.AreEqual "pragma_function_list", .TableName, "TableName mismatch."
        Assert.AreEqual "rowid", .Name, "Name mismatch."
        Assert.AreEqual "rowid", .OriginName, "Name mismatch."
    End With
    
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


'@TestMethod("Metadata")
Private Sub ztcColumnMetaAPI_ThrowsOnUninitializedSQLiteColumnMeta()
    On Error Resume Next
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsPragmaTable
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
Act:
    Dim ColumnInfo As SQLiteCColumnMeta
    ColumnInfo.ColumnIndex = 1
    '''' Throws if this not set: ColumnInfo.Initialized = -1
    ResultCode = dbs.DbExecutor.ColumnMetaAPI(ColumnInfo)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Metadata")
Private Sub ztcGetTableMeta_VerifiesFunctionsTableMeta()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObj.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsTable
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbExecutor.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbs.DbExecutor.TableMeta
Assert:
    Assert.AreEqual 0, LBound(TableMeta), "TableMeta base mismatch."
    Assert.AreEqual 5, UBound(TableMeta), "TableMeta size mismatch."
    Assert.AreEqual "enc", TableMeta(3).Name, "enc column name mismatch."
    Assert.AreEqual "narg", TableMeta(4).Name, "nargs column name mismatch "
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


'@TestMethod("Metadata")
Private Sub ztcGetTableMeta_ThrowsOnUnpreparedStatement()
    On Error Resume Next
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.CREATETableITRBrowid
    Dim AffectedRows As Long
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error."
Act:
    ResultCode = dbs.DbExecutor.TableMetaCollect

    Guard.AssertExpectedError Assert, StatementNotPreparedErr
End Sub


'@TestMethod("Metadata")
Private Sub ztcGetTableMeta_VerifiesFunctionsTableMetaRowid()
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
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObj.CreateFunctionsTableWithData(dbc)
'    Assert.IsTrue AffectedRows > 0, "Failed to INSERT test data."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsTableRowid
    ResultCode = dbs.Prepare16V2(SQLQuery)
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbExecutor.TableMetaCollect
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbs.DbExecutor.TableMeta
Assert:
    Assert.AreEqual 0, LBound(TableMeta), "TableMeta base mismatch."
    Assert.AreEqual 6, UBound(TableMeta), "TableMeta size mismatch."
    Assert.AreEqual "enc", TableMeta(4).Name, "enc column name mismatch."
    Assert.AreEqual "narg", TableMeta(5).Name, "nargs column name mismatch "
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

''@TestMethod("Metadata")
'Private Sub ztcGetTableMeta_VerifiesSQLiteErrorWithUnpreparedStatement()
'    On Error GoTo TestFail
'
'    Set FixObj = New SQLiteCTestFixObj
'    Set FixSQL = New SQLiteCTestFixSQL
'Arrange:
'    Dim dbc As SQLiteCConnection
'    Set dbc = FixObj.GetConnDbMemory
'    Dim dbs As SQLiteCStatement
'    Set dbs = dbc.CreateStatement(vbNullString)
'
'    Dim ResultCode As SQLiteResultCodes
'    ResultCode = dbc.OpenDb
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
'    Dim AffectedRows As Long
'Act:
'    Dim SQLQuery As String
'    SQLQuery = FixSQL.CREATETableITRBrowid
'    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
'    ResultCode = dbs.DbExecutor.GetTableMeta
'Assert:
'    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error (GetTableMeta)."
'Cleanup:
'    ResultCode = dbs.Finalize
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
'    ResultCode = dbc.CloseDb
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
'
'CleanExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
'End Sub
'
'
''@TestMethod("Metadata")
'Private Sub ztcGetTableMeta_ThrowsOnUnpreparedStatement()
'    On Error Resume Next
'
'Arrange:
'    Dim dbc As SQLiteCConnection
'    Set dbc = FixObj.GetConnDbMemory
'    Dim dbs As SQLiteCStatement
'    Set dbs = dbc.CreateStatement(vbNullString)
'
'    Dim ResultCode As SQLiteResultCodes
'
'    Dim SQLQuery As String
'    SQLQuery = FixSQL.FunctionsPragmaTable
'    ResultCode = dbc.OpenDb
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
'    ResultCode = dbs.Prepare16V2(SQLQuery)
'    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
'Act:
'    Dim ColumnInfo As SQLiteCColumnMeta
'    ColumnInfo.ColumnIndex = 1
'    '''' Throws if this not set: ColumnInfo.Initialized = -1
'    ResultCode = dbs.DbExecutor.GetColumnMetaAPI(ColumnInfo)
'
'    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
'End Sub
'
