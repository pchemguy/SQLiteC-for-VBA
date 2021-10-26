Attribute VB_Name = "SQLiteCExecSQLTesting"
'@Folder "SQLite.SQLiteC For VBA.Statement"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
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


'@TestMethod("Data types")
Private Sub ztcSQLiteTypeName_VerifiesSQLiteTypeName()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
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
    Set dbm = FixMain.ObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
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
    Set dbm = FixMain.ObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
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
    Set dbm = FixMain.ObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
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
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLMain.Func.SelectPragmaRowid
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
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLMain.Func.SelectPragmaNoRowid
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
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixMain.ObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLMain.Func.SelectNoRowid
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
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLMain.ITRB.CreateRowid
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

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    Dim SQLQuery As String
    Dim AffectedRows As Long

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    SQLQuery = FixSQLMain.ITRB.CreateRowidWithValues
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.IsTrue AffectedRows = 5, "Failed to INSERT test data."
Act:
    SQLQuery = FixSQLMain.ITRB.SelectRowid
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbExecutor.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbs.DbExecutor.TableMeta
Assert:
    Assert.AreEqual 0, LBound(TableMeta), "TableMeta base mismatch."
    Assert.AreEqual 5, UBound(TableMeta), "TableMeta size mismatch."
    Assert.AreEqual "xr", TableMeta(4).Name, "enc column name mismatch."
    Assert.AreEqual "xb", TableMeta(5).Name, "nargs column name mismatch."
    
    With TableMeta(0)
        Assert.AreEqual "rowid", .Name, "rowid column name mismatch."
        Assert.IsTrue .RowId, "Rowid should be true."
        Assert.IsTrue .PrimaryKey, "PrimaryKey should be true."
        Assert.IsFalse .AutoIncrement, "AutoIncrement should be false."
        Assert.IsFalse .NotNull, "AutoIncrement should be false."
        Assert.AreEqual SQLITE_AFF_INTEGER, .Affinity, "Expected Affinity=SQLITE_AFF_INTEGER"
        Assert.AreEqual SQLITE_INTEGER, .AffinityType, "Expected AffinityType=SQLITE_INTEGER"
        Assert.AreEqual "INTEGER", .DeclaredTypeT, "Expected DeclaredTypeT=INTEGER"
    End With
    
    With TableMeta(1)
        Assert.AreEqual "id", .Name, "id column name mismatch."
        Assert.IsFalse .RowId, "Rowid should be false."
        Assert.IsTrue .PrimaryKey, "PrimaryKey should be true."
        Assert.IsFalse .AutoIncrement, "AutoIncrement should be false."
        Assert.IsTrue .NotNull, "AutoIncrement should be true."
        Assert.AreEqual SQLITE_AFF_INTEGER, .Affinity, "Expected Affinity=SQLITE_AFF_INTEGER"
        Assert.AreEqual SQLITE_INTEGER, .AffinityType, "Expected AffinityType=SQLITE_INTEGER"
        Assert.AreEqual "INT", .DeclaredTypeT, "Expected DeclaredTypeT=INT"
        Assert.AreEqual "BINARY", .Collation, "Expected Collation=BINARY"
        Assert.AreEqual "main", .DbName, "Expected DbName=main"
        Assert.AreEqual "itrb", .TableName, "Expected TableName=itrb"
    End With
    
    With TableMeta(3)
        Assert.AreEqual SQLITE_AFF_TEXT, .Affinity, "Expected Affinity=SQLITE_AFF_TEXT"
        Assert.AreEqual SQLITE_TEXT, .AffinityType, "Expected AffinityType=SQLITE_TEXT"
        Assert.AreEqual "TEXT", .DeclaredTypeT, "Expected DeclaredTypeT=TEXT"
        Assert.AreEqual "NOCASE", .Collation, "Expected Collation=NOCASE"
    End With
    
    With TableMeta(4)
        Assert.AreEqual SQLITE_AFF_REAL, .Affinity, "Expected Affinity=SQLITE_AFF_REAL"
        Assert.AreEqual SQLITE_FLOAT, .AffinityType, "Expected AffinityType=SQLITE_FLOAT"
        Assert.AreEqual "REAL", .DeclaredTypeT, "Expected DeclaredTypeT=REAL"
        Assert.IsTrue .NotNull, "AutoIncrement should be true."
    End With
    
    With TableMeta(5)
        Assert.AreEqual SQLITE_AFF_BLOB, .Affinity, "Expected Affinity=SQLITE_AFF_BLOB"
        Assert.AreEqual SQLITE_BLOB, .AffinityType, "Expected AffinityType=SQLITE_BLOB"
        Assert.AreEqual "BLOB", .DeclaredTypeT, "Expected DeclaredTypeT=BLOB"
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
