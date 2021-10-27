Attribute VB_Name = "SQLiteCTests"
'@Folder "SQLite.SQLiteC.Manager"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
Option Explicit
Option Private Module

Private Const LITE_LIB As String = "SQLiteCDBVBA"
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
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfo()
    On Error GoTo TestFail

Arrange:
    Dim DllPath As String
    Dim DllNames As Variant
    #If Win64 Then
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
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfoV2()
    On Error GoTo TestFail

Arrange:
    Dim DllPath As String
    #If Win64 Then
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
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesDefaultManager()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM
Assert:
    Assert.IsNotNothing dbm, "Default manager is not set."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcGetMainDbId_VerifiesIsNull()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM
Assert:
    Assert.IsTrue IsNull(dbm.MainDbId), "Main db is not null."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcGetDllMan_VerifiesIsSet()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM
Assert:
    Assert.IsNotNothing dbm.DllMan, "Dll manager is not set"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnMan")
Private Sub ztcConnDb_VerifiesIsNotSet()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM
Assert:
    Assert.IsNothing dbm.ConnDb(vbNullString), "Connection should be nothing"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsGivenWrongDllBitness()
    On Error Resume Next
    Dim DllPath As String
    Dim DllNames As Variant
    #If Win64 Then
        DllPath = LITE_RPREFIX & "dll\x32"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = LITE_RPREFIX & "dll\x64"
        DllNames = "sqlite3.dll"
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath, DllNames)
    Guard.AssertExpectedError Assert, LoadingDllErr
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsOnInvalidDllPath()
    On Error Resume Next
    Dim DllPath As String
    DllPath = "____INVALID PATH____"
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath)
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("Connection")
Private Sub ztcCreateConnection_VerifiesSQLiteCConnectionWithValidDbPath()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCReg
Assert:
    Assert.IsNotNothing dbc, "Default SQLiteCConnection is not set."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbConn_VerifiesSavedConnectionReference()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM()
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
Assert:
    Assert.IsNotNothing DbConn, "Default SQLiteCConnection is not set."
    Assert.AreEqual DbPathName, dbm.MainDbId
    Assert.AreSame DbConn, dbm.ConnDb(DbPathName)
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbConn_VerifiesMemoryMainDb()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM()
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
Assert:
    Assert.AreEqual DbPathName, dbm.MainDbId
    Assert.AreSame DbConn, dbm.ConnDb(DbPathName)
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection")
Private Sub ztcGetDbConn_VerifiesTempMainDb()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixMain.ObjC.GetDBM()
    Dim DbPathName As String
    DbPathName = ":blank:"
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
Assert:
    Assert.AreEqual vbNullString, dbm.MainDbId
    Assert.AreSame DbConn, dbm.ConnDb(vbNullString)
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
