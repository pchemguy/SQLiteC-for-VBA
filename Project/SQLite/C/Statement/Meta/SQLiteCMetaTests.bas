Attribute VB_Name = "SQLiteCMetaTests"
'@Folder "SQLite.C.Statement.Meta"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCMetaTests"
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


'@TestMethod("Data types")
Private Sub ztcSQLiteTypeName_VerifiesSQLiteTypeName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
Act:
Assert:
    Assert.AreEqual "INTEGER", dbsm.SQLiteTypeName(SQLITE_INTEGER), "SQLiteTypeName mismatch."
    Assert.AreEqual "FLOAT", dbsm.SQLiteTypeName(SQLITE_FLOAT), "SQLiteTypeName mismatch."
    Assert.AreEqual "TEXT", dbsm.SQLiteTypeName(SQLITE_TEXT), "SQLiteTypeName mismatch."
    Assert.AreEqual "NULL", dbsm.SQLiteTypeName(SQLITE_NULL), "SQLiteTypeName mismatch."
    Assert.AreEqual "BLOB", dbsm.SQLiteTypeName(SQLITE_BLOB), "SQLiteTypeName mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcSQLiteTypeAffinityName_VerifiesSQLiteTypeAffinityName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
Act:
Assert:
    Assert.AreEqual "BLOB", dbsm.SQLiteTypeAffinityName(SQLITE_AFF_BLOB), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "INTEGER", dbsm.SQLiteTypeAffinityName(SQLITE_AFF_INTEGER), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "NUMERIC", dbsm.SQLiteTypeAffinityName(SQLITE_AFF_NUMERIC), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "REAL", dbsm.SQLiteTypeAffinityName(SQLITE_AFF_REAL), "SQLiteTypeAffinityName mismatch."
    Assert.AreEqual "TEXT", dbsm.SQLiteTypeAffinityName(SQLITE_AFF_TEXT), "SQLiteTypeAffinityName mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcTypeAffinityFromDeclaredType_VerifiesDeclaredTypeHandling()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
Act:
Assert:
    Assert.AreEqual SQLITE_AFF_INTEGER, dbsm.TypeAffinityFromDeclaredType("UNSIGNED BIG iNt"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_TEXT, dbsm.TypeAffinityFromDeclaredType("NATIVE cHaRACTER(70)"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_BLOB, dbsm.TypeAffinityFromDeclaredType("BLoB"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_REAL, dbsm.TypeAffinityFromDeclaredType("DOuBLE PRECISION"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_NUMERIC, dbsm.TypeAffinityFromDeclaredType("STRING"), "TypeAffinityFromDeclaredType mismatch."
    Assert.AreEqual SQLITE_AFF_INTEGER, dbsm.TypeAffinityFromDeclaredType("FLoATING POInT"), "TypeAffinityFromDeclaredType mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Data types")
Private Sub ztcTypeAffinityMap_VerifiesMappingToSQLiteTypes()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
Act:
Assert:
    Assert.AreEqual SQLITE_NONE, dbsm.AffinityMap(SQLITE_AFF_NONE - SQLITE_AFF_NONE), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_BLOB, dbsm.AffinityMap(SQLITE_AFF_BLOB - SQLITE_AFF_NONE), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_TEXT, dbsm.AffinityMap(SQLITE_AFF_TEXT - SQLITE_AFF_NONE), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_TEXT, dbsm.AffinityMap(SQLITE_AFF_NUMERIC - SQLITE_AFF_NONE), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_INTEGER, dbsm.AffinityMap(SQLITE_AFF_INTEGER - SQLITE_AFF_NONE), "AffinityMap mismatch."
    Assert.AreEqual SQLITE_FLOAT, dbsm.AffinityMap(SQLITE_AFF_REAL - SQLITE_AFF_NONE), "AffinityMap mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Metadata")
Private Sub ztcColumnMetaAPI_VerifiesFunctionsColumnMeta()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectPragmaRowid
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
    ResultCode = dbsm.ColumnMetaAPI(ColumnInfo)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Unexpected GetColumnMetaAPI error."
Assert:
    With ColumnInfo
        Assert.AreEqual SQLITE_AFF_INTEGER, .Affinity, "Affinity mismatch."
        Assert.AreEqual "main", .DbName, "Db alias mismatch."
        Assert.AreEqual "pragma_function_list", .TableName, "TableName mismatch."
        Assert.AreEqual "rowid", .Name, "Name mismatch."
        Assert.AreEqual "rowid", .OriginName, "Name mismatch."
    End With
    
CleanUp:
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
    TestCounter = TestCounter + 1
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectPragmaNoRowid
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
Act:
    Dim ColumnInfo As SQLiteCColumnMeta
    ColumnInfo.ColumnIndex = 1
    '''' Throws if this not set: ColumnInfo.Initialized = -1
    ResultCode = dbsm.ColumnMetaAPI(ColumnInfo)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub











'@TestMethod("Metadata")
Private Sub ztcTableMetaCollect_VerifiesFunctionsTableMeta()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectNoRowid
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbsm.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbsm.TableMeta
Assert:
    Assert.AreEqual 0, LBound(TableMeta), "TableMeta base mismatch."
    Assert.AreEqual 5, UBound(TableMeta), "TableMeta size mismatch."
    Assert.AreEqual "enc", TableMeta(3).Name, "enc column name mismatch."
    Assert.AreEqual "narg", TableMeta(4).Name, "nargs column name mismatch "
CleanUp:
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
Private Sub ztcTableMetaCollect_ThrowsOnUnpreparedStatement()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
    
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateRowid
    Dim AffectedRows As Long
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error."
Act:
    ResultCode = dbsm.TableMetaCollect

    Guard.AssertExpectedError Assert, StatementNotPreparedErr
End Sub


'@TestMethod("Metadata")
Private Sub ztcTableMetaCollect_VerifiesFunctionsTableMetaRowid()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)

    Dim ResultCode As SQLiteResultCodes
    Dim SQLQuery As String
    Dim AffectedRows As Long

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    SQLQuery = FixSQLITRB.CreateRowidWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.IsTrue AffectedRows = 5, "Failed to INSERT test data."
Act:
    SQLQuery = FixSQLITRB.SelectRowid
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbsm.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbsm.TableMeta
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
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
