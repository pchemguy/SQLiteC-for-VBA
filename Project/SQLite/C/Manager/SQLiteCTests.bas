Attribute VB_Name = "SQLiteCTests"
'@Folder "SQLite.C.Manager"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext, StopKeyword
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCTests"
Private TestCounter As Long
'''' When >0, test runner will stop at every test, starting from #STOP_IN_TEST
Private Const STOP_IN_TEST As Long = 0

Private Const LITE_LIB As String = "SQLiteCAdo"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

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


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfo()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If
    
Arrange:
    Dim DllPath As String
    Dim DllNames As Variant
    #If WIN64 Then
        DllPath = LITE_RPREFIX & "dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = LITE_RPREFIX & "dll\x32"
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath, DllNames)
Act:
    Dim VersionS As String
    VersionS = Replace(dbm.Version(False), ".", "0") & "0"
    Dim VersionN As String
    VersionN = CStr(dbm.Version(True))
Assert:
    Assert.AreEqual VersionS, VersionN, "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfoV2()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim DllPath As String
    #If WIN64 Then
        DllPath = LITE_RPREFIX & "dll\x64"
    #Else
        DllPath = LITE_RPREFIX & "dll\x32"
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath)
Act:
    Dim VersionS As String
    VersionS = Replace(dbm.Version(False), ".", "0") & "0"
    Dim VersionN As String
    VersionN = CStr(dbm.Version(True))
Assert:
    Assert.AreEqual VersionS, VersionN, "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesDefaultManager()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
Assert:
    Assert.IsNotNothing dbm, "Default manager is not set."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("Factory")
Private Sub ztcGetMainDbId_VerifiesIsNull()
    Exit Sub
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(vbNullString)

Assert:
    Assert.IsTrue IsNull(dbm.MainDbId), "Main db is not null."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("Factory")
Private Sub ztcGetDllMan_VerifiesIsSet()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
Assert:
    Assert.IsNotNothing dbm.DllMan, "Dll manager is not set"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("ConnMan")
Private Sub ztcConnDb_VerifiesIsNotSet()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
Assert:
    Assert.IsNothing dbm.ConnDb(vbNullString), "Connection should be nothing"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsGivenWrongDllBitness()
    On Error Resume Next
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If
    Dim DllPath As String
    Dim DllNames As Variant
    #If WIN64 Then
        DllPath = LITE_RPREFIX & "dll\x32"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = LITE_RPREFIX & "dll\x64"
        DllNames = "sqlite3.dll"
    #End If
    DllManager.ForgetSingleton
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath, DllNames)
    Guard.AssertExpectedError Assert, LoadingDllErr
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsOnInvalidDllPath()
    On Error Resume Next
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If
    Dim DllPath As String
    DllPath = "____INVALID PATH____"
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath)
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'''' Crashes Excel on exit when run with other tests. Works fine when run alone
''@TestMethod("Connection")
'Private Sub ztcCreateConnection_VerifiesSQLiteCConnectionWithValidDbPath()
'    On Error GoTo TestFail
'    TestCounter = TestCounter + 1
'    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
'        Debug.Print "TEST #: " & CStr(TestCounter)
'        Stop
'    End If
'
'Arrange:
'    Dim dbc As SQLiteCConnection
'    Set dbc = FixObjC.GetDBCReg
'Assert:
'    Assert.IsNotNothing dbc, "Default SQLiteCConnection is not set."
'
'CleanExit:
'    Exit Sub
'TestFail:
'    If Not Assert Is Nothing Then
'        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
'    Else
'        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
'    End If
'End Sub


'''' Crashes Excel on exit when run with other tests. Works fine when run alone
''@TestMethod("Connection")
'Private Sub ztcGetDbConn_VerifiesSavedConnectionReference()
'    On Error GoTo TestFail
'    TestCounter = TestCounter + 1
'    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
'        Debug.Print "TEST #: " & CStr(TestCounter)
'        Stop
'    End If
'
'Arrange:
'    Dim dbm As SQLiteC
'    Set dbm = FixObjC.GetDBM()
'    Dim DbPathName As String
'    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
'    Dim DbConn As SQLiteCConnection
'    Set DbConn = dbm.CreateConnection(DbPathName)
'Assert:
'    Assert.IsNotNothing DbConn, "Default SQLiteCConnection is not set."
'    Assert.AreEqual DbPathName, dbm.MainDbId, "dbm.MainDbId mismatch"
'    Assert.AreSame DbConn, dbm.ConnDb(DbPathName), "Connection reference mismatch"
'
'CleanExit:
'    Exit Sub
'TestFail:
'    If Not Assert Is Nothing Then
'        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
'    Else
'        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
'    End If
'End Sub


'''' Crashes Excel on exit when run with other tests. Works fine when run alone
''@TestMethod("Connection")
'Private Sub ztcGetDbConn_VerifiesMemoryMainDb()
'    On Error GoTo TestFail
'    TestCounter = TestCounter + 1
'    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
'        Debug.Print "TEST #: " & CStr(TestCounter)
'        Stop
'    End If
'
'Arrange:
'    '''' In general, tests may reuse the db manager. This test verifies the
'    '''' state of the freshly instantiated SQLiteC object, so cleanup must
'    '''' be executed before the test. (dbm.MainDbId is set the first time
'    '''' the dbm.CreateConnection is called.)
'    Dim dbm As SQLiteC
'    Set dbm = FixObjC.GetDBM()
'    Dim DbPathName As String
'    DbPathName = ":memory:"
'    Dim DbConn As SQLiteCConnection
'    Set DbConn = dbm.CreateConnection(DbPathName)
'Assert:
'    Assert.AreEqual DbPathName, dbm.MainDbId, "dbm.MainDbId mismatch"
'    Assert.AreSame DbConn, dbm.ConnDb(DbPathName), "Connection reference mismatch"
'
'CleanExit:
'    Exit Sub
'TestFail:
'    If Not Assert Is Nothing Then
'        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
'    Else
'        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
'    End If
'End Sub


'''' Crashes Excel on exit when run with other tests. Works fine when run alone
''@TestMethod("Connection")
'Private Sub ztcGetDbConn_VerifiesTempMainDb()
'    On Error GoTo TestFail
'    TestCounter = TestCounter + 1
'    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
'        Debug.Print "TEST #: " & CStr(TestCounter)
'        Stop
'    End If
'
'Arrange:
'    '''' In general, tests may reuse the db manager. This test verifies the
'    '''' state of the freshly instantiated SQLiteC object, so cleanup must
'    '''' be executed before the test. (dbm.MainDbId is set the first time
'    '''' the dbm.CreateConnection is called.)
'    Dim dbm As SQLiteC
'    Set dbm = FixObjC.GetDBM()
'    Dim DbPathName As String
'    DbPathName = ":blank:"
'    Dim DbConn As SQLiteCConnection
'    Set DbConn = dbm.CreateConnection(DbPathName)
'Assert:
'    Assert.AreEqual vbNullString, dbm.MainDbId, "dbm.MainDbId mismatch"
'    Assert.AreSame DbConn, dbm.ConnDb(vbNullString), "Connection reference mismatch"
'
'CleanExit:
'    Exit Sub
'TestFail:
'    If Not Assert Is Nothing Then
'        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
'    Else
'        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
'    End If
'End Sub


'@TestMethod("Backup")
Private Sub ztcDupDbOnlineFull_VerifiesDbCopyMemToTemp()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbcSrc As SQLiteCConnection
    Set dbcSrc = FixObjC.GetDBCMemITRBWithData
    FixObjC.ReuseDBM
    Dim dbcDst As SQLiteCConnection
    Set dbcDst = FixObjC.GetDBCTmp

    Assert.AreEqual SQLITE_OK, dbcSrc.OpenDb, "Unexpected OpenDb error."
    Assert.AreEqual SQLITE_OK, dbcDst.OpenDb, "Unexpected OpenDb error."

    Dim DbStmtNameSrc As String
    DbStmtNameSrc = Left$(GenerateGUID, 8)
    Dim dbsSrc As SQLiteCStatement
    Set dbsSrc = dbcSrc.CreateStatement(DbStmtNameSrc)
    Dim DbStmtNameDst As String
    DbStmtNameDst = Left$(GenerateGUID, 8)
    Dim dbsDst As SQLiteCStatement
    Set dbsDst = dbcDst.CreateStatement(DbStmtNameDst)

    Dim SQLQuery As String
    SQLQuery = "SELECT count(*) As counter FROM itrb"
    Assert.AreEqual 5, dbsSrc.GetScalar(SQLQuery), "Unexpected RowCount."
Act:
    Dim PagesDone As Long
    PagesDone = SQLiteC.DupDbOnlineFull(dbcDst, "main", dbcSrc, "main")
Assert:
    Assert.AreEqual 3, PagesDone, "PagesDone mismatch."
    Assert.AreEqual 5, dbsDst.GetScalar(SQLQuery), "Unexpected RowCount."
CleanUp:
    Assert.AreEqual SQLITE_OK, dbcSrc.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbcDst.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


''@TestMethod("Backup")
Private Sub ztcDupDbOnlineFull_VerifiesDbCopyTempToMem()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    If STOP_IN_TEST <> 0 And STOP_IN_TEST <= TestCounter Then
        Debug.Print "TEST #: " & CStr(TestCounter)
        Stop
    End If

Arrange:
    Dim dbcSrc As SQLiteCConnection
    Set dbcSrc = FixObjC.GetDBCTmpITRBWithData
    Dim dbcDst As SQLiteCConnection
    Set dbcDst = FixObjC.GetDBCMem

    Assert.AreEqual SQLITE_OK, dbcSrc.OpenDb, "Unexpected OpenDb error."
    Assert.AreEqual SQLITE_OK, dbcDst.OpenDb, "Unexpected OpenDb error."

    Dim DbStmtNameSrc As String
    DbStmtNameSrc = Left$(GenerateGUID, 8)
    Dim dbsSrc As SQLiteCStatement
    Set dbsSrc = dbcSrc.CreateStatement(DbStmtNameSrc)
    Dim DbStmtNameDst As String
    DbStmtNameDst = Left$(GenerateGUID, 8)
    Dim dbsDst As SQLiteCStatement
    Set dbsDst = dbcDst.CreateStatement(DbStmtNameDst)

    Dim SQLQuery As String
    SQLQuery = "SELECT count(*) As counter FROM itrb"
    Assert.AreEqual 5, dbsSrc.GetScalar(SQLQuery), "Unexpected RowCount."
Act:
    Dim PagesDone As Long
    PagesDone = SQLiteC.DupDbOnlineFull(dbcDst, "main", dbcSrc, "main")
Assert:
    Assert.AreEqual 3, PagesDone, "PagesDone mismatch."
    Assert.AreEqual 5, dbsDst.GetScalar(SQLQuery), "Unexpected RowCount."
CleanUp:
    Assert.AreEqual SQLITE_OK, dbcSrc.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbcDst.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub
