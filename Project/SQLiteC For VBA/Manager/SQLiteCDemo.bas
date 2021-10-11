Attribute VB_Name = "SQLiteCDemo"
'@Folder "SQLiteC For VBA.Manager"
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
    Dim DbConn As SQLiteCConnection
    Set DbConn = ConnFix.ConnDbRegular
    Debug.Print DbConn.Version(False)
    Debug.Print CStr(DbConn.Version(True))
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
