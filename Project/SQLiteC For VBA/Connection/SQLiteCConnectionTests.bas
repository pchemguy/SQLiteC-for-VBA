Attribute VB_Name = "SQLiteCConnectionTests"
'@Folder "SQLiteC For VBA.Connection"
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

Private Type TSQLiteCConnectionTests
    DllMan As DllManager
End Type
Private this As TSQLiteCConnectionTests


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    
    Dim DllPath As String
    Dim DllNames As Variant
    #If WIN64 Then
        DllPath = "Library\SQLiteCforVBA\dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = "Library\SQLiteCforVBA\dll\x32"
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Set this.DllMan = DllManager.Create(DllPath, DllNames)
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set this.DllMan = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("SQLiteVersion")
Private Sub ztcSQLite3Version_VerifiesVersionInfo()
    On Error GoTo TestFail

Arrange:
    If this.DllMan Is Nothing Then
        Debug.Print "Loading SQLite in ztcSQLite3Version_VerifiesVersionInfo"
        Dim DllPath As String
        Dim DllNames As Variant
        #If WIN64 Then
            DllPath = "Library\SQLiteCforVBA\dll\x64"
            DllNames = "sqlite3.dll"
        #Else
            DllPath = "Library\SQLiteCforVBA\dll\x32"
            DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
        Set this.DllMan = DllManager.Create(DllPath, DllNames)
    End If
Act:
    Dim DbConn As SQLiteCConnection
    Set DbConn = SQLiteCConnection(vbNullString)
    Dim VersionS As String
    VersionS = Replace(DbConn.Version(False), ".", "0") & "0"
    Dim VersionN As String
    VersionN = CStr(DbConn.Version(True))
    Set this.DllMan = Nothing
Assert:
    Assert.AreEqual VersionS, VersionN, "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
