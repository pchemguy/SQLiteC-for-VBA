Attribute VB_Name = "SQLiteCConnectionOpenCloseTests"
'@Folder "SQLite.SQLiteC.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
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


'@TestMethod("Connection")
Private Sub ztcCreateConnection_VerifiesSQLiteCConnectionWithValidDbPath()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCReg
Assert:
    Assert.IsNotNothing dbc, "Default SQLiteCConnection is not set."
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"
    Assert.IsNotNothing dbc.ErrorInfo, "ErrorInfo must be set."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbPathName_VerifiesMemoryDbPathName()
    On Error GoTo TestFail

Arrange:
Act:
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
Assert:
    Assert.AreEqual DbPathName, dbc.DbPathName
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbPathName_VerifiesAnonDbPathName()
    On Error GoTo TestFail

Arrange:
Act:
    Dim DbPathName As String
    DbPathName = vbNullString
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCAnon
Assert:
    Assert.AreEqual DbPathName, dbc.DbPathName
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcAttachedDbPathName_ThrowsOnClosedConnection()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
    Debug.Print dbc.DbPathName = dbc.AttachedDbPathName
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("Connection")
Private Sub ztcAttachedDbPathName_VerifiesTempDbPathName()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCTemp
    Dim ResultCode As SQLiteResultCodes
Act:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    Assert.AreEqual dbc.DbPathName, dbc.AttachedDbPathName, "AttachedDbPathName mismatch."
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcAtDetachDatabase_VerifiesAttachExistingNewMem()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTemp
    Dim dbcTemp As SQLiteCConnection
    Set dbcTemp = FixObjC.GetDBCTemp
    
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbcTemp.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    ResultCode = dbcTemp.ExecuteNonQueryPlain(FixSQLGeneral.CreateBasicTable)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    Assert.AreEqual SQLITE_OK, dbcTemp.CloseDb, "Unexpected CloseDb error"
    Dim NewDbPathName As String
    NewDbPathName = FixObjC.RandomTempFileName
    
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"

    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(DbStmtName)
    Assert.IsNotNothing dbs, "Failed to create an SQLiteCStatement instance."
    
    Dim fso As New Scripting.FileSystemObject
    Dim SQLDbCount As String
    SQLDbCount = "SELECT count(*) As counter FROM pragma_database_list"
Act:
Assert:
    ResultCode = dbc.AttachDatabase(dbcTemp.DbPathName)
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach existing db"
    Assert.AreEqual 2, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (exist)."
    ResultCode = dbc.AttachDatabase(":memory:")
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach memory db"
    Assert.AreEqual 3, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (memory)."
    ResultCode = dbc.AttachDatabase(NewDbPathName)
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach new db"
    Assert.AreEqual 4, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (new)."
    
    ResultCode = dbc.DetachDatabase(fso.GetBaseName(NewDbPathName))
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach new db"
    Assert.AreEqual 3, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (new)."
    ResultCode = dbc.DetachDatabase("memory")
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach memory db"
    Assert.AreEqual 2, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (memory)."
    ResultCode = dbc.DetachDatabase(fso.GetBaseName(dbcTemp.DbPathName))
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach existing db"
    Assert.AreEqual 1, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (exist)."
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithRegularDb()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCReg
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithTempDb()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCAnon
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbConnection")
Private Sub ztcOpenDbCloseDb_VerifiesWithMemoryDb()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMem
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
    Assert.AreNotEqual 0, dbc.DbHandle, "DbHandle must not be 0"
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    Assert.AreEqual 0, dbc.DbHandle, "DbHandle must be 0"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
