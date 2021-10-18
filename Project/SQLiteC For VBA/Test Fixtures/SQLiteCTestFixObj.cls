VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCTestFixObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides objects for tests"
'@Folder "SQLiteC For VBA.Test Fixtures"
'@ModuleDescription "Provides objects for tests"
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
    zfxGetConnDbMemoryWithTable.ExecuteNonQueryPlain SQLiteCTestFixSQL.CREATETableITRB
End Function


