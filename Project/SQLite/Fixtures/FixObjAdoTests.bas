Attribute VB_Name = "FixObjAdoTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "FixObjAdoTests"
Private TestCounter As Long

#Const LateBind = 1     '''' RubberDuck Tests
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
Private Sub ztcGetDbReg_VerifiesDbReg()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbReg()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left$(dbq.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    Assert.AreEqual FixObjAdo.DefaultDbPathName, dbq.MainDB, "Database name (dbq.MainDB) mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbq.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "_" & dbq.MainDB & "_", dbq.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbAnon_VerifiesDbAnon()
    On Error GoTo TestFail

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbAnon()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left$(dbq.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbq.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "__", dbq.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbMem_VerifiesDbMem()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMem()
    Dim SQLiteMajorVersion As String
    SQLiteMajorVersion = "3"
    Dim SQLQuery As String
Act:
Assert:
    SQLQuery = "SELECT sqlite_version()"
    Assert.AreEqual SQLiteMajorVersion, Left$(dbq.GetScalar(SQLQuery), 1), "SQLiteMajorVersion mismatch."
    SQLQuery = "SELECT count(*) FROM pragma_database_list()"
    Assert.AreEqual 1, dbq.GetScalar(SQLQuery), "Database count mismatch."
    SQLQuery = "SELECT '_' || file || '_' FROM pragma_database_list() WHERE name='main'"
    Assert.AreEqual "__", dbq.GetScalar(SQLQuery), "Database name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbMemITRB_VerifiesDbMemITRB()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMemITRB
Act:
    Dim TableDDLExpected As String
    TableDDLExpected = FixSQLITRB.Create
    Dim SQLQuery As String
    SQLQuery = "SELECT [sql] || ';' FROM" & _
               SQLlib.SubQuery(LiteMetaSQL.Create("main").Tables) & _
               "WHERE tbl_name = 'itrb'"
    Dim TableDDLActual As String
    TableDDLActual = dbq.GetScalar(SQLQuery)
    SQLQuery = SQLlib.CountSelect(LiteMetaSQL.Create("main").Tables)
Assert:
    Assert.AreEqual 1, dbq.GetScalar(SQLQuery), "Table count mismatch."
    Assert.AreEqual TableDDLExpected, TableDDLActual, "Table CREATE mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbMemITRBWithData_VerifiesDbMemITRBWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMemITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbq.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbMemFuncWithData_VerifiesDbMemFuncWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMemFuncWithData
    Dim Expected As Long
    Expected = dbq.GetScalar("SELECT count(*) FROM pragma_function_list()")
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbq.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbTmp_VerifiesDbTmp()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbTmp
    Dim TmpPath As String
    TmpPath = FixObjAdo.RandomTempFileName
Act:
Assert:
    Assert.AreEqual Len(TmpPath), Len(dbq.MainDB), "TMP database path template is wrong."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbTmpITRBWithData_VerifiesGetDBMTmpITRBWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbTmpITRBWithData
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLITRB.SelectNoRowid)
Assert:
    Assert.AreEqual 5, dbq.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDbTmpFuncWithData_VerifiesDbTmpFuncWithData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbTmpFuncWithData
    Dim Expected As Long
    Expected = dbq.GetScalar("SELECT count(*) FROM pragma_function_list()")
Act:
    Dim SQLQuery As String
    SQLQuery = SQLlib.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    Assert.AreEqual Expected, dbq.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


