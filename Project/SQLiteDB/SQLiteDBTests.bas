Attribute VB_Name = "SQLiteDBTests"
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
Private Sub ztcCreate_ValidatesDefaultPath()
    On Error GoTo TestFail

Arrange:
    Dim DatabaseName As String
    DatabaseName = Left$(ThisWorkbook.Name, InStr(Len(ThisWorkbook.Name) - 5, ThisWorkbook.Name, ".xl")) & "db"
    Dim ExpectedPath As String
    ExpectedPath = ThisWorkbook.Path & Application.PathSeparator & DatabaseName
    Dim ExpectedConnectionString As String
    ExpectedConnectionString = "Driver=SQLite3 ODBC Driver;" + _
                               "Database=" + ThisWorkbook.Path + Application.PathSeparator + DatabaseName + ";" + _
                               "SyncPragma=NORMAL;FKSupport=True;"
Act:
    DatabaseName = vbNullString
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB.Create(DatabaseName)
    Dim ActualPath As String
    ActualPath = DbManager.MainDB
    Dim ActualConnectionString As String
    ActualConnectionString = DbManager.ConnectionString
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = DbManager.GetAdoRecordset
Assert:
    Assert.AreEqual ExpectedPath, ActualPath, "Default path mismatch"
    Assert.AreEqual ExpectedConnectionString, ActualConnectionString, "Default ConnectionString mismatch"
    Assert.IsNotNothing DbManager.AdoCommand, "AdoCommand is not set"
    Assert.IsNotNothing DbManager.AdoConnection, "AdoConnection is not set"
    Assert.AreEqual "SELECT sqlite_version() AS version", DbManager.AdoCommand.CommandText, "AdoCommand SQL mismatch"
    Assert.IsNotNothing AdoRecordset, "Default AdoRecordset is not set"
    Assert.IsNothing AdoRecordset.ActiveConnection, "Default AdoRecordset is not disconnected"
    Assert.AreEqual 1, AdoRecordset.RecordCount, "Expected one record"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

