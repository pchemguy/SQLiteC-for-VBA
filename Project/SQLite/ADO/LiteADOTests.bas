Attribute VB_Name = "LiteADOTests"
'@Folder "SQLite.ADO"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration: it's ok for services (FileSystemObject)
Option Explicit
Option Private Module

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

Private Const PATH_SEP As String = "\"


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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesExistingDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = FixObjAdo.DefaultDbPathName
Act:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMReg()
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "Existing db path mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcFromConnection_ValidatesNewDbManager()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = LiteADO(FixObjAdo.DefaultDbPathName)
    Dim dbmCI As LiteADO
    Set dbmCI = dbm
Act:
    Dim dbmClone As ILiteADO
    Set dbmClone = LiteADO.FromConnection(dbmCI.AdoConnection)
    Dim dbmCloneCI As LiteADO
    Set dbmCloneCI = dbm
Assert:
    Assert.AreEqual dbm.MainDB, dbmClone.MainDB, "Db path mismatch"
    Assert.IsTrue dbmCI.AdoConnection Is dbmCloneCI.AdoConnection, "Bad connection"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesInMemoryDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ":memory:"
Act:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMem()
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "InMemory path mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesNewRelativeDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim FileName As String
    FileName = "NewDB" & GenerateGUID & ".tmp"
    Dim RelativePathName As String
    RelativePathName = "Temp" & PATH_SEP & FileName
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & RelativePathName
    '''' This test creates a new db file that remains locked for a certain
    '''' period of time. If this test is rerun too soon, deletion will fail.
    On Error Resume Next
    MkDir ThisWorkbook.Path & PATH_SEP & "Temp"
    Kill ThisWorkbook.Path & PATH_SEP & "Temp" & PATH_SEP & "*.tmp"
    On Error GoTo TestFail
Act:
    Dim dbm As ILiteADO
    Set dbm = LiteADO.Create(RelativePathName, AllowNonExistent:=True)
    Dim dbmCI As LiteADO
    Set dbmCI = dbm
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
CleanUp:
    dbmCI.AdoConnection.Close
    Set dbm = Nothing
    Set dbmCI = Nothing

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesNewAbsoluteDatabasePath()
    On Error GoTo TestFail
    
Arrange:
    Dim Expected As String
    Expected = FixObjAdo.RandomTempFileName
Act:
    Dim dbm As ILiteADO
    Set dbm = LiteADO.Create(Expected, AllowNonExistent:=True)
    Dim dbmCI As LiteADO
    Set dbmCI = dbm
    Dim Actual As String
    Actual = dbm.MainDB
CleanUp:
    dbmCI.AdoConnection.Close
    Set dbm = Nothing
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesDefaultConnectionString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;NoCreat=True;"
Act:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMReg()
    Dim Actual As String
    Actual = dbm.ConnectionString
Assert:
    Assert.AreEqual Expected, Actual, "Default ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesNoCreatConnectionString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;NoCreat=True;"
Act:
    Dim dbm As ILiteADO
    Set dbm = LiteADO(FixObjAdo.DefaultDbPathName, False)
    Dim Actual As String
    Actual = dbm.ConnectionString
Assert:
    Assert.AreEqual Expected, Actual, "Default ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesCreatConnectionString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim dbm As ILiteADO
    Set dbm = LiteADO(FixObjAdo.DefaultDbPathName, True)
    Dim Actual As String
    Actual = dbm.ConnectionString
Assert:
    Assert.AreEqual Expected, Actual, "Default ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Recordset")
Private Sub ztcCreate_ValidatesDefaultRecordset()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMReg()
    Dim dbmCI As LiteADO
    Set dbmCI = dbm
    Dim DefaultSQL As String
    DefaultSQL = "SELECT sqlite_version() AS version"
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = dbm.GetAdoRecordset(vbNullString)
Assert:
    Assert.IsNotNothing dbmCI.AdoCommand, "AdoCommand is not set"
    Assert.IsNotNothing dbmCI.AdoConnection, "AdoConnection is not set"
    Assert.AreEqual DefaultSQL, dbmCI.AdoCommand.CommandText, "SQL mismatch"
    Assert.IsNotNothing AdoRecordset, "AdoRecordset is not set"
    Assert.IsNothing AdoRecordset.ActiveConnection, "AdoRecordset is not disconnected"
    Assert.AreEqual 1, AdoRecordset.RecordCount, "Expected record count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
