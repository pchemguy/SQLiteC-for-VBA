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
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
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
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
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
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
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
    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
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
