Attribute VB_Name = "SQLiteCDemo"
'@Folder "SQLiteC For VBA.Manager"
'@IgnoreModule ProcedureNotUsed
Option Explicit
Option Private Module


'''' Custom functions added to SQLite source for testing/verification purposes
#If VBA7 Then
Private Declare PtrSafe Function sqlite3_latin_utf8 Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_cyrillic_utf8 Lib "SQLite3" () As LongPtr ' PtrUtf8String
#Else
Private Declare Function sqlite3_latin_utf8 Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_cyrillic_utf8 Lib "SQLite3" () As Long ' PtrUtf8String
#End If


Private Sub GetSQLiteVersionString()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    '@Ignore VariableNotUsed
    Dim dbm As SQLiteC
    '@Ignore AssignmentNotUsed
    Set dbm = ConnFix.dbm
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbRegular
    Debug.Print DbConn.Version(False)
    Debug.Print CStr(DbConn.Version(True))
    
    '''' This test functions are only available in a custom built SQLite library
    On Error GoTo FUNCTION_NOT_AVAILABLE:
    Debug.Print CStr(ConnFix.LibVersionNumber)
    Debug.Print ConnFix.LatinUTF8
    Debug.Print ConnFix.CyrillicUTF8
    Debug.Print CStr(DbConn.VersionI64)
    
    On Error GoTo 0
    Exit Sub
    
FUNCTION_NOT_AVAILABLE:
    Const DllFunctionNotFoundErr As Long = 453
    Const ErrorMessage As String = "Can't find DLL entry point sqlite3_libversion_number_i64"
    If Err.Number = DllFunctionNotFoundErr And _
       Left(Err.Description, Len(ErrorMessage)) = ErrorMessage Then
        MsgBox "sqlite3_libversion_number_i64 is only available in a custom built SQLite library!"
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub


Private Sub OpenCloseDbRegular()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbRegular
    DbConn.OpenDb
    DbConn.CloseDb
End Sub


Private Sub OpenCloseDbInvalidPath()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbInvalidPath
    DbConn.OpenDb
    DbConn.CloseDb
End Sub


Private Sub OpenCloseDbNotDb()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbNotDb
    DbConn.OpenDb
    DbConn.CloseDb
End Sub


Private Sub OpenCloseLockedDb()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbLockedDb
    DbConn.OpenDb
    DbConn.CloseDb
End Sub


Private Sub TestDbRegular()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim dbm As SQLiteC
    Set dbm = ConnFix.dbm
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbRegular
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = ConnFix.StmtDb("main")
    
    Dim Result As Variant
    Result = DbConn.Version
    Result = DbConn.Version(False)
    Result = DbConn.VersionI64
    DbConn.OpenDb
    Result = DbStmt.GetScalar("SELECT sqlite_version()")
    Result = DbStmt.GetRowSet("SELECT * FROM pragma_module_list()")
    DbConn.CloseDb
End Sub
