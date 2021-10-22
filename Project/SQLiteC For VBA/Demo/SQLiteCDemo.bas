Attribute VB_Name = "SQLiteCDemo"
'@Folder "SQLiteC For VBA.Demo"
'@IgnoreModule ProcedureNotUsed
Option Explicit
Option Private Module


Private Sub GetSQLiteVersionString()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    '@Ignore VariableNotUsed
    Dim dbm As SQLiteC
    '@Ignore AssignmentNotUsed
    Set dbm = ConnFix.dbm
    Debug.Print dbm.Version(False)
    Debug.Print CStr(dbm.Version(True))
    
    '''' This test functions are only available in a custom built SQLite library
    On Error GoTo FUNCTION_NOT_AVAILABLE:
    Debug.Print CStr(ConnFix.LibVersionNumber)
    Debug.Print ConnFix.LatinUTF8
    Debug.Print ConnFix.CyrillicUTF8
    Debug.Print CStr(ConnFix.VersionI64)
    
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
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbConn.OpenDb
    Debug.Assert ResultCode = SQLITE_OK
    ResultCode = DbConn.CloseDb
    Debug.Assert ResultCode = SQLITE_OK
End Sub


Private Sub OpenCloseDbInvalidPath()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbInvalidPath
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbConn.OpenDb
    Debug.Assert ResultCode = SQLITE_OK
    ResultCode = DbConn.CloseDb
    Debug.Assert ResultCode = SQLITE_OK
End Sub


Private Sub OpenCloseDbNotDb()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbNotDb
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbConn.OpenDb
    Debug.Assert ResultCode = SQLITE_OK
    ResultCode = DbConn.CloseDb
    Debug.Assert ResultCode = SQLITE_OK
End Sub


Private Sub OpenCloseLockedDb()
    Dim ConnFix As SQLiteCConnDemoFix
    Set ConnFix = SQLiteCConnDemoFix.Create
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbLockedDb
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbConn.OpenDb
    Debug.Assert ResultCode = SQLITE_OK
    ResultCode = DbConn.CloseDb
    Debug.Assert ResultCode = SQLITE_OK
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
    Result = dbm.Version
    Debug.Print Result
    Result = dbm.Version(False)
    Debug.Print Result
    Result = ConnFix.VersionI64
    Debug.Print Result
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbConn.OpenDb
    Debug.Assert ResultCode = SQLITE_OK
    Result = DbStmt.GetScalar("SELECT sqlite_version()")
    Debug.Print Result
    Result = DbStmt.GetPagedRowSet("SELECT * FROM pragma_module_list()")
    Debug.Print Result(0)(0)(0)
    ResultCode = DbConn.CloseDb
    Debug.Assert ResultCode = SQLITE_OK
End Sub

