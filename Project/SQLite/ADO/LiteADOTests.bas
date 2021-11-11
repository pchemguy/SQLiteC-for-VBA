Attribute VB_Name = "LiteADOTests"
'@Folder "SQLite.ADO"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration: it's ok for services (FileSystemObject)
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LiteADOTests"
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
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesExistingDatabasePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = FixObjAdo.DefaultDbPathName
Act:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbReg()
    Dim Actual As String
    Actual = dbq.MainDB
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
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = LiteADO(FixObjAdo.DefaultDbPathName)
    Dim dbqCI As LiteADO
    Set dbqCI = dbq
Act:
    Dim dbqClone As ILiteADO
    Set dbqClone = LiteADO.FromConnection(dbqCI.AdoConnection)
    Dim dbqCloneCI As LiteADO
    Set dbqCloneCI = dbq
Assert:
    Assert.AreEqual dbq.MainDB, dbqClone.MainDB, "Db path mismatch"
    Assert.IsTrue dbqCI.AdoConnection Is dbqCloneCI.AdoConnection, "Bad connection"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesInMemoryDatabasePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = ":memory:"
Act:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMem()
    Dim Actual As String
    Actual = dbq.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "InMemory path mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ValidatesNewAbsoluteDatabasePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    
Arrange:
    Dim Expected As String
    Expected = FixObjAdo.RandomTempFileName
Act:
    Dim dbq As ILiteADO
    Set dbq = LiteADO.Create(Expected, AllowNonExistent:=True)
    Dim dbqCI As LiteADO
    Set dbqCI = dbq
    Dim Actual As String
    Actual = dbq.MainDB
Cleanup:
    dbqCI.AdoConnection.Close
    Set dbq = Nothing
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
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;NoCreat=True;"
Act:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbReg()
    Dim Actual As String
    Actual = dbq.ConnectionString
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
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;NoCreat=True;"
Act:
    Dim dbq As ILiteADO
    Set dbq = LiteADO(FixObjAdo.DefaultDbPathName, False)
    Dim Actual As String
    Actual = dbq.ConnectionString
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
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & FixObjAdo.DefaultDbPathName & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim dbq As ILiteADO
    Set dbq = LiteADO(FixObjAdo.DefaultDbPathName, True)
    Dim Actual As String
    Actual = dbq.ConnectionString
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
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbReg()
    Dim dbqCI As LiteADO
    Set dbqCI = dbq
    Dim DefaultSQL As String
    DefaultSQL = "SELECT sqlite_version() AS version"
Act:
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = dbq.GetAdoRecordset(vbNullString)
Assert:
    Assert.IsNotNothing dbqCI.AdoCommand, "AdoCommand is not set"
    Assert.IsNotNothing dbqCI.AdoConnection, "AdoConnection is not set"
    Assert.AreEqual DefaultSQL, dbqCI.AdoCommand.CommandText, "SQL mismatch"
    Assert.IsNotNothing AdoRecordset, "AdoRecordset is not set"
    Assert.IsNothing AdoRecordset.ActiveConnection, "AdoRecordset is not disconnected"
    Assert.AreEqual 1, AdoRecordset.RecordCount, "Expected record count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
