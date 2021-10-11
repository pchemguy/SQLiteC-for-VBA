Attribute VB_Name = "SQLiteCTests"
'@Folder "SQLiteC For VBA.Manager"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
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


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfo()
    On Error GoTo TestFail

Arrange:
    Dim DllPath As String
    Dim DllNames As Variant
    #If WIN64 Then
        DllPath = "Library\SQLiteCforVBA\dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = "Library\SQLiteCforVBA\dll\x32"
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath, DllNames)
Act:
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.DbConnInit(vbNullString)
    Dim VersionS As String
    VersionS = Replace(DbConn.Version(False), ".", "0") & "0"
    Dim VersionN As String
    VersionN = CStr(DbConn.Version(True))
Assert:
    Assert.AreEqual VersionS, VersionN, "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQLiteVersion")
'@Ignore UseMeaningfulName
Private Sub ztcSQLite3Version_VerifiesVersionInfoV2()
    On Error GoTo TestFail

Arrange:
    Dim DllPath As String
    #If WIN64 Then
        DllPath = "Library\SQLiteCforVBA\dll\x64"
    #Else
        DllPath = "Library\SQLiteCforVBA\dll\x32"
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath)
Act:
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.DbConnInit(vbNullString)
    Dim VersionS As String
    VersionS = Replace(DbConn.Version(False), ".", "0") & "0"
    Dim VersionN As String
    VersionN = CStr(DbConn.Version(True))
Assert:
    Assert.AreEqual VersionS, VersionN, "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

