Attribute VB_Name = "MinADOTests"
'@Folder "SQLiteDB"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

Private Const LIB_NAME As String = "SQLiteDBVBA"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


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
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxDefaultDbManager() As MinADO
    Dim FilePathName As String
    FilePathName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim dbm As MinADO
    Set dbm = MinADO.Create(FilePathName)
    Set zfxDefaultDbManager = dbm
End Function


Private Function zfxMemoryDbManager() As MinADO
    Set zfxMemoryDbManager = MinADO.Create(":memory:")
End Function


Private Function zfxDefaultDbPath() As String
    zfxDefaultDbPath = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesExistingDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & LIB_NAME & ".db"
Act:
    Dim dbm As MinADO
    Set dbm = zfxDefaultDbManager()
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
Private Sub ztcCreate_ValidatesInMemoryDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ":memory:"
Act:
    Dim dbm As MinADO
    Set dbm = zfxMemoryDbManager()
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
    Dim RelativePathName As String
    RelativePathName = REL_PREFIX & "NewDB.sqlite"
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & RelativePathName
Act:
    Dim dbm As MinADO
    Set dbm = MinADO.Create(RelativePathName, AllowNonExistent:=True)
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
    
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
    Expected = ThisWorkbook.Path & PATH_SEP & "NewDB.sqlite"
Act:
    Dim dbm As MinADO
    Set dbm = MinADO.Create(Expected, AllowNonExistent:=True)
    Dim Actual As String
    Actual = dbm.MainDB
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
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & zfxDefaultDbPath & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim dbm As MinADO
    Set dbm = zfxDefaultDbManager()
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
    Dim dbm As MinADO
    Set dbm = zfxDefaultDbManager()
    Dim DefaultSQL As String
    DefaultSQL = "SELECT sqlite_version() AS version"
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = dbm.GetAdoRecordset
Assert:
    Assert.IsNotNothing dbm.AdoCommand, "AdoCommand is not set"
    Assert.IsNotNothing dbm.AdoConnection, "AdoConnection is not set"
    Assert.AreEqual DefaultSQL, dbm.AdoCommand.CommandText, "SQL mismatch"
    Assert.IsNotNothing AdoRecordset, "AdoRecordset is not set"
    Assert.IsNothing AdoRecordset.ActiveConnection, "AdoRecordset is not disconnected"
    Assert.AreEqual 1, AdoRecordset.RecordCount, "Expected record count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
