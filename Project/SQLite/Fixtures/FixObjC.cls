VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixObjC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides instantiated library objects"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides instantiated SQLiteC library objects"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Private Const LITE_LIB As String = "SQLiteCDBVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP
'@Ignore SelfAssignedDeclaration
Private fso As New Scripting.FileSystemObject

Private Type TFixObjC
    DllMan As DllManager
    LibDir As String
End Type
Private this As TFixObjC


Private Sub Class_Initialize()
    this.LibDir = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX
End Sub

Public Function GetDBM() As SQLiteC
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim dbm As SQLiteC
    '''' Using default library names hardcoded in the SQLiteC constructor.
    Set dbm = SQLiteC(DllPath)
    If dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create an SQLiteC instance."
    Set GetDBM = dbm
End Function

Public Function GetDBC(ByVal DbPathName As String) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = GetDBM()
    '''' Keep a reference to the DllManager to prevent library unloading due to object destruction
    Set this.DllMan = dbm.DllMan
    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(DbPathName)
    If dbc Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create an SQLiteCConnection instance."
    Set GetDBC = dbc
End Function

Public Function GetDBCDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.LibDir & LITE_LIB & ".db"
    Set GetDBCDbRegular = GetDBC(DbPathName)
End Function

Public Function GetDBCDbReadOnlyAttr() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.LibDir & "ReadOnlyAttr.db"
    With fso.GetFile(DbPathName)
        .Attributes = .Attributes Or ReadOnly
    End With
    Set GetDBCDbReadOnlyAttr = GetDBC(DbPathName)
End Function

Public Function GetDBCDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetDBCDbMemory = GetDBC(DbPathName)
End Function

Public Function GetDBCDbAnon() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set GetDBCDbAnon = GetDBC(DbPathName)
End Function

Public Function GetDBCDbTemp() As SQLiteCConnection
    Dim DbFilePath As String
    DbFilePath = Environ("TEMP")
    Dim DbFileName As String
    DbFileName = Replace(Replace(Replace(Now(), "/", "-"), " ", "_"), ":", "-")
    Dim DbFileExt As String
    DbFileExt = ".db"
    Dim DbPathName As String
    DbPathName = fso.BuildPath(DbFilePath, DbFileName) & DbFileExt
    Set GetDBCDbTemp = GetDBC(DbPathName)
End Function

Public Function GetDBCInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set GetDBCInvalidPath = GetDBC(DbPathName)
End Function

Public Function GetDBCDbMemoryWithITRB() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetDBCDbMemoryWithITRB = GetDBC(DbPathName)
    With GetDBCDbMemoryWithITRB
        .OpenDb
        .ExecuteNonQueryPlain FixSQLITRB.Create
        .CloseDb
    End With
End Function

Public Function CreateFunctionsTableWithData(ByVal dbc As SQLiteCConnection) As Long
    Dim SQLQuery As String
    SQLQuery = FixSQLFunctions.CreateWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    CreateFunctionsTableWithData = AffectedRows
End Function
