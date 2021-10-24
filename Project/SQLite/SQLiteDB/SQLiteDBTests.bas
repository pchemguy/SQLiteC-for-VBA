Attribute VB_Name = "SQLiteDBTests"
'@Folder "SQLite.SQLiteDB"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

Private Const LIB_NAME As String = "SQLiteCDBVBA"
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


Private Function zfxDefaultDbManager() As SQLiteDB
    Dim FilePathName As String
    FilePathName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim dbm As SQLiteDB
    Set dbm = SQLiteDB.Create(FilePathName)
    Set zfxDefaultDbManager = dbm
End Function


Private Function zfxMemoryDbManager() As SQLiteDB
    Set zfxMemoryDbManager = SQLiteDB.Create(":memory:")
End Function


'@EntryPoint
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
    Dim dbm As SQLiteDB
    Set dbm = zfxDefaultDbManager()
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "Existing db path mismatch"
Cleanup:
    dbm.AdoConnection.Close
    Set dbm = Nothing
    
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
    Dim dbm As SQLiteDB
    Set dbm = zfxMemoryDbManager()
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "InMemory path mismatch"
Cleanup:
    dbm.AdoConnection.Close
    Set dbm = Nothing
    
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
    '''' period of time. If this test is rerun to soon, deletion will fail.
    On Error Resume Next
    MkDir ThisWorkbook.Path & PATH_SEP & "Temp"
    Kill ThisWorkbook.Path & PATH_SEP & "Temp" & PATH_SEP & "*.tmp"
    On Error GoTo TestFail
Act:
    Dim dbm As SQLiteDB
    Set dbm = SQLiteDB.Create(RelativePathName, AllowNonExistent:=True)
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
Cleanup:
    dbm.AdoConnection.Close
    Set dbm = Nothing
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesNewAbsoluteDatabasePath()
    On Error GoTo TestFail

Arrange:
    Dim FileName As String
    FileName = "NewDB" & GenerateGUID & ".tmp"
    Dim RelativePathName As String
    RelativePathName = "Temp" & PATH_SEP & FileName
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & RelativePathName
    '''' This test creates a new db file that remains locked for a certain
    '''' period of time. If this test is rerun to soon, deletion will fail.
    On Error Resume Next
    MkDir ThisWorkbook.Path & PATH_SEP & "Temp"
    Kill ThisWorkbook.Path & PATH_SEP & "Temp" & PATH_SEP & "*.tmp"
    On Error GoTo TestFail
Act:
    Dim dbm As SQLiteDB
    Set dbm = SQLiteDB.Create(Expected, AllowNonExistent:=True)
    Dim Actual As String
    Actual = dbm.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
Cleanup:
    dbm.AdoConnection.Close
    Set dbm = Nothing
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
