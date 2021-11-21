Attribute VB_Name = "LiteManTests"
'@Folder "SQLite.ADO"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration: it's ok for services (FileSystemObject)
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LiteManTests"
Private TestCounter As Long

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


'@TestMethod("Environment")
Private Sub ztcSQLite3ODBCDriverCheck_VerifiesSQLiteODBCDriver()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
Assert:
    Assert.IsTrue LiteMan.SQLite3ODBCDriverCheck(), "SQLite3ODBC driver not found."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path Resolution")
Private Sub ztcCreate_ValidatesNewRelativeDatabasePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = "NewDB" & GenerateGUID & ".tmp"
    Dim RelativePathName As String
    RelativePathName = "Temp" & PATH_SEP & FileName
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & RelativePathName
    '''' This test creates a new db file that may remain locked for a certain
    '''' period of time. If this test is rerun too soon, deletion may fail.
    ' On Error Resume Next
    Dim Prefix As String
    Prefix = ThisWorkbook.Path & PATH_SEP & "Temp"
    If fso.FolderExists(Prefix) Then
        fso.DeleteFile Prefix & PATH_SEP & "*.*"
    Else
        fso.CreateFolder Prefix
    End If
    'On Error GoTo TestFail
Act:
    Dim dbm As LiteMan
    Set dbm = LiteMan(RelativePathName, AllowNonExistent:=True)
    Dim dbq As ILiteADO
    Set dbq = dbm.ExecADO
    Dim Actual As String
    Actual = dbq.MainDB
Assert:
    Assert.AreEqual Expected, Actual, "New db (relative) path mismatch"
CleanUp:
    Set dbq = Nothing
    Set dbm = Nothing
    On Error Resume Next
    fso.DeleteFolder Prefix
    On Error GoTo TestFail

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
