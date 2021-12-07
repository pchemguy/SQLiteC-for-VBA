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
CleanUp:
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
               ";" & LiteADO.DefaultOptions & "NoCreat=True;"
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
               ";" & LiteADO.DefaultOptions & "NoCreat=True;"
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
               ";" & LiteADO.DefaultOptions
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


'@TestMethod("Options")
Private Sub ztcODBCOptionsStr_ValidatesDefaultODBCOptionsStr()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Expected = LiteADO.DefaultOptions
Act:
    Dim Actual As String
    Actual = LiteADO.ODBCOptionsStr
Assert:
    Assert.AreEqual Expected, Actual, "Default options mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Options")
Private Sub ztcDefaultOptionsDict_ValidatesDefaultOptionsDictType()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
Assert:
    Assert.IsTrue IsObject(LiteADO.DefaultOptionsDict), "Default options dict type mismatch"
    Assert.AreEqual "Dictionary", TypeName(LiteADO.DefaultOptionsDict), "Default options dict type mismatch"
    Assert.AreEqual 5, LiteADO.DefaultOptionsDict.Count, "Default options dict count mismatch"
    Assert.AreEqual True, LiteADO.DefaultOptionsDict("FKSupport"), "Default option FKSupport mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Options")
Private Sub ztcODBCOptionsStr_ValidatesEmptyOptions()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbqCI As LiteADO
    Set dbqCI = LiteADO(vbNullString, False, Empty)
Assert:
    Assert.AreEqual 6, dbqCI.ODBCOptions.Count, "ODBC options count mismatch."
    Assert.AreEqual True, dbqCI.ODBCOptions("FKSupport"), "ODBC option FKSupport mismatch."
    Assert.AreEqual True, dbqCI.ODBCOptions("NoCreat"), "ODBC option NoCreat mismatch."
    Assert.AreEqual dbqCI.DefaultOptions & "NoCreat=True;", dbqCI.ODBCOptionsStr, "ODBC options str mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Options")
Private Sub ztcODBCOptionsStr_ValidatesStringOptions()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbqCI As LiteADO
    Set dbqCI = LiteADO(vbNullString, True, "FKSupport=False;")
Assert:
    Assert.IsTrue vbString = VarType(dbqCI.ODBCOptions), "ODBC options type mismatch."
    Assert.AreEqual dbqCI.ODBCOptionsStr, dbqCI.ODBCOptions, "ODBC options mismatch."
    Assert.AreEqual "FKSupport=False;", dbqCI.ODBCOptions, "ODBC options mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Options")
Private Sub ztcODBCOptionsStr_ValidatesDictOptions()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Options As Scripting.Dictionary
    Set Options = New Scripting.Dictionary
    Options("FKSupport") = False
    Options("Timeout") = 50000
    Dim Expected As String
    Expected = Replace(LiteADO.DefaultOptions, "FKSupport=True;", "FKSupport=False;") _
               & "Timeout=50000;NoCreat=True;"
Act:
    Dim dbqCI As LiteADO
    Set dbqCI = LiteADO(vbNullString, False, Options)
Assert:
    Assert.IsTrue IsObject(dbqCI.ODBCOptions), "ODBC options type mismatch."
    Assert.AreEqual "Dictionary", TypeName(dbqCI.ODBCOptions), "ODBC options type mismatch."
    Assert.AreEqual 7, dbqCI.ODBCOptions.Count, "ODBC options count mismatch."
    Assert.AreEqual False, dbqCI.ODBCOptions("FKSupport"), "ODBC option FKSupport mismatch."
    Assert.AreEqual True, dbqCI.ODBCOptions("NoCreat"), "ODBC option NoCreat mismatch."
    Assert.AreEqual 50000, dbqCI.ODBCOptions("Timeout"), "ODBC option Timeout mismatch."
    Assert.AreEqual Expected, dbqCI.ODBCOptionsStr, "ODBC options mismatch."
    
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
    Assert.IsFalse dbqCI.AdoCommand Is Nothing, "AdoCommand is not set"
    Assert.IsFalse dbqCI.AdoConnection Is Nothing, "AdoConnection is not set"
    Assert.AreEqual DefaultSQL, dbqCI.AdoCommand.CommandText, "SQL mismatch"
    Assert.IsFalse AdoRecordset Is Nothing, "AdoRecordset is not set"
    Assert.IsTrue AdoRecordset.ActiveConnection Is Nothing, "AdoRecordset is not disconnected"
    Assert.AreEqual 1, AdoRecordset.RecordCount, "Expected record count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
