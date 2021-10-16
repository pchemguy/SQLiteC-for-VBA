VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCTestFixtures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SQLiteC For VBA.Test Fixtures"
'@PredeclaredId
'@IgnoreModule MoveFieldCloserToUsage, ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Private Const LITE_LIB As String = "SQLiteCforVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TSQLiteCConnectionTestFixtures
    DllMan As DllManager
End Type
Private this As TSQLiteCConnectionTestFixtures


Public Function zfxGetDefaultDBM() As SQLiteC
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim dbm As SQLiteC
    '''' Using default library names hardcoded in the SQLiteC constructor.
    Set dbm = SQLiteC(DllPath)
    If dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteC instance."
    Set zfxGetDefaultDBM = dbm
End Function

Public Function zfxGetConnection(ByVal DbPathName As String) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = zfxGetDefaultDBM()
    '''' Keep a reference to the DllManager to prevent library unloading due to object destruction
    Set this.DllMan = dbm.DllMan
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteCConnection instance."
    Set zfxGetConnection = DbConn
End Function

Public Function zfxGetConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set zfxGetConnDbRegular = zfxGetConnection(DbPathName)
End Function

Public Function zfxGetConnDbReadOnlyAttr() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & "ReadOnlyAttr.db"
    Set zfxGetConnDbReadOnlyAttr = zfxGetConnection(DbPathName)
End Function

Public Function zfxGetConnDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set zfxGetConnDbMemory = zfxGetConnection(DbPathName)
End Function

Public Function zfxGetConnDbTemp() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set zfxGetConnDbTemp = zfxGetConnection(DbPathName)
End Function

Public Function zfxGetConnDbInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set zfxGetConnDbInvalidPath = zfxGetConnection(DbPathName)
End Function

Public Function zfxGetConnDbMemoryWithTable() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set zfxGetConnDbMemoryWithTable = zfxGetConnection(DbPathName)
    zfxGetConnDbMemoryWithTable.ExecuteNonQueryPlain zfxGetStmtCreateTableIRBNT2V
End Function

Public Function zfxGetStmtCreateTableIRBNT() As String
    Const TEST_STMT_CREATE_TABLE_IRBNT As String = _
        "DROP TABLE IF EXISTS t1;" & vbNewLine & _
        "CREATE TABLE t1(" & vbNewLine & _
        "    xi INTEGER," & vbNewLine & _
        "    xr REAL," & vbNewLine & _
        "    xb BLOB," & vbNewLine & _
        "    xn NUMERIC," & vbNewLine & _
        "    xt Text" & vbNewLine & _
        ");" & vbNewLine & _
        "INSERT INTO t1(rowid, xi, xr, xb, xn,    xt) " & vbNewLine & _
        "VALUES        (    1, 10, .1,  1,  7, 'AAA')," & vbNewLine & _
        "              (    2, 20, .3,  1, .2, 'BBB')," & vbNewLine & _
        "              (    3,  8, .2,  1, 10, 'AAA')," & vbNewLine & _
        "              (    4, 27, .3,  1, .9, 'DDD')," & vbNewLine & _
        "              (    5,  2, .8,  1,  1, 'BBB');" & vbNewLine
    zfxGetStmtCreateTableIRBNT = TEST_STMT_CREATE_TABLE_IRBNT
End Function

Public Function zfxGetStmtCreateTableIRBNT2V() As String
    Const TEST_STMT_CREATE_TABLE_IRBNT As String = _
        "CREATE TABLE t1(" & vbNewLine & _
        "    xi INTEGER," & vbNewLine & _
        "    xr REAL," & vbNewLine & _
        "    xb BLOB," & vbNewLine & _
        "    xn NUMERIC," & vbNewLine & _
        "    xt Text" & vbNewLine & _
        ");" & vbNewLine
    zfxGetStmtCreateTableIRBNT2V = TEST_STMT_CREATE_TABLE_IRBNT
End Function

Public Function zfxGetStmtSQLiteVersion() As String
    zfxGetStmtSQLiteVersion = "SELECT sqlite_version()"
End Function

Public Function zfxGetStmtSELECTt1() As String
    zfxGetStmtSELECTt1 = "SELECT * FROM t1"
End Function

