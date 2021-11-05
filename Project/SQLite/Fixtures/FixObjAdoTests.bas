Attribute VB_Name = "FixObjAdoTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
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


'@TestMethod("Fixture")
Private Sub ztcGetDBMReg_VerifiesDBMReg()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMReg()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left(dbm.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    Assert.AreEqual FixObjAdo.DefaultDbPathName, dbm.MainDB, "Database name (dbm.MainDB) mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbm.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "_" & dbm.MainDB & "_", dbm.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMAnon_VerifiesDBMAnon()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMAnon()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left(dbm.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbm.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "__", dbm.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMMem_VerifiesDBMMem()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMem()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left(dbm.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbm.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "__", dbm.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMMemITRB_VerifiesDBMMemITRB()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMemITRB
Act:
    Dim TableDDLExpected As String
    TableDDLExpected = FixSQLITRB.Create
    Dim SQLQuery As String
    SQLQuery = "SELECT [sql] || ';' FROM" & _
               FixSQLMisc.SubQuery(LiteMetaSQL("main").Tables) & _
               "WHERE tbl_name = 'itrb'"
    Dim TableDDLActual As String
    TableDDLActual = dbm.GetScalar(SQLQuery)
    SQLQuery = FixSQLMisc.CountSelect(LiteMetaSQL("main").Tables)
Assert:
    Assert.AreEqual 1, dbm.GetScalar(SQLQuery), "Table count mismatch."
    Assert.AreEqual TableDDLExpected, TableDDLActual, "Table CREATE mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMMemITRBWithData_VerifiesDBMMemITRBWithData()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMemITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLMisc.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbm.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMMemFuncWithData_VerifiesDBMMemFuncWithData()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMemFuncWithData
    Dim Expected As Long
    Expected = dbm.GetScalar("SELECT count(*) FROM pragma_function_list()")
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLMisc.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbm.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMTmp_VerifiesDBMTmp()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMTmp
    Dim TmpPath As String
    TmpPath = FixObjAdo.RandomTempFileName
Act:
Assert:
    Assert.AreEqual Len(TmpPath), Len(dbm.MainDB), "TMP database path template is wrong."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMTmpITRBWithData_VerifiesGetDBMTmpITRBWithData()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMTmpITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLMisc.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbm.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBMTmpFuncWithData_VerifiesDBMTmpFuncWithData()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMTmpFuncWithData
    Dim Expected As Long
    Expected = dbm.GetScalar("SELECT count(*) FROM pragma_function_list()")
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLMisc.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbm.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
