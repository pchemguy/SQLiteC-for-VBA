VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixObjC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides instantiated SQLiteC library objects"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides instantiated SQLiteC library objects"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Private Const LITE_LIB As String = "SQLiteCDBVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP
'@Ignore SelfAssignedDeclaration
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

Public Function GetDBC(ByVal DbPathName As String, _
              Optional ByVal AllowNonExistent As Variant = True) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = GetDBM()
    '''' Keep a reference to the DllManager to prevent library unloading due to object destruction
    Set this.DllMan = dbm.DllMan
    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(DbPathName, AllowNonExistent)
    If dbc Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create an SQLiteCConnection instance."
    Set GetDBC = dbc
End Function

Public Function GetDBCReg() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.LibDir & LITE_LIB & ".db"
    Set GetDBCReg = GetDBC(DbPathName)
End Function

Public Function GetDBCReadOnlyAttr() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.LibDir & "ReadOnlyAttr.db"
    Dim fso As New Scripting.FileSystemObject
    With fso.GetFile(DbPathName)
        .Attributes = .Attributes Or ReadOnly
    End With
    Set GetDBCReadOnlyAttr = GetDBC(DbPathName)
End Function

Public Function GetDBCAnon() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":blank:"
    Set GetDBCAnon = GetDBC(DbPathName)
End Function

Public Function GetDBCInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set GetDBCInvalidPath = GetDBC(DbPathName)
End Function

Public Function GetDBCMem() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetDBCMem = GetDBC(DbPathName)
End Function

'''' Returns opened in-memory database with "ITRB" table
Public Function GetDBCMemITRB() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetDBCMemITRB = GetDBC(DbPathName)
    With GetDBCMemITRB
        If .OpenDb <> SQLITE_OK Then
            Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to open db table."
        End If
        .ExecuteNonQueryPlain FixSQLITRB.Create
        '''' DO NOT CLOSE In-Memory Database!
    End With
End Function

'''' Returns opened in-memory database with "functions" table and data
Public Function GetDBCMemFuncWithData() As SQLiteCConnection
    Dim AllowNonExistent As Boolean
    AllowNonExistent = True
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBC(DbPathName, AllowNonExistent)
    If dbc.OpenDb <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to open db table."
    End If
    Dim AffectedRows As Long
    AffectedRows = CreateFunctionsTableWithData(dbc)
    If AffectedRows <= 0 Then Set dbc = Nothing
    '''' DO NOT CLOSE In-Memory Database!
    Set GetDBCMemFuncWithData = dbc
End Function

'''' Returns opened in-memory database with "itrb" table and data
Public Function GetDBCMemITRBWithData() As SQLiteCConnection
    Dim AllowNonExistent As Boolean
    AllowNonExistent = True
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBC(DbPathName, AllowNonExistent)
    If dbc.OpenDb <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to open db table."
    End If
    Dim AffectedRows As Long
    AffectedRows = CreateITRBTableWithData(dbc)
    If AffectedRows <= 0 Then Set dbc = Nothing
    '''' DO NOT CLOSE In-Memory Database!
    Set GetDBCMemITRBWithData = dbc
End Function

Public Function CreateFunctionsTableWithData(ByVal dbc As SQLiteCConnection) As Long
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    Dim SQLQuery As String
    SQLQuery = FixSQLFunctions.CreateWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    Dim ResultCode As SQLiteResultCodes
    If HandleConnection Then
        If dbc.OpenDb <> SQLITE_OK Then
            Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to open db table."
        End If
    End If
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    If HandleConnection Then
        If dbc.CloseDb <> SQLITE_OK Then
            Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to close db table."
        End If
    End If
    CreateFunctionsTableWithData = AffectedRows
End Function

Public Function CreateITRBTableWithData(ByVal dbc As SQLiteCConnection) As Long
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    Dim ResultCode As SQLiteResultCodes
    If HandleConnection Then
        If dbc.OpenDb <> SQLITE_OK Then
            Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to open db table."
        End If
    End If
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    If HandleConnection Then
        If dbc.CloseDb <> SQLITE_OK Then
            Err.Raise ErrNo.UnknownClassErr, "FixObjC", "Failed to close db table."
        End If
    End If
    CreateITRBTableWithData = AffectedRows
End Function

Public Function RandomTempFileName(Optional ByVal FileExt As String = ".db")
    RandomTempFileName = Environ$("TEMP") & PATH_SEP & _
        Format$(Now, "yyyy_mm_dd-hh_mm_ss") & "-" & Left$(GenerateGUID, 8) & FileExt
End Function

Public Function GetDBCTemp() As SQLiteCConnection
    Dim AllowNonExistent As Boolean
    AllowNonExistent = True
    Dim DbPathName As String
    DbPathName = RandomTempFileName()
    Set GetDBCTemp = GetDBC(DbPathName, AllowNonExistent)
End Function

Public Function GetDBCTempFuncWithData() As SQLiteCConnection
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTemp()
    Dim AffectedRows As Long
    AffectedRows = CreateFunctionsTableWithData(dbc)
    If AffectedRows <= 0 Then Set dbc = Nothing
    Set GetDBCTempFuncWithData = dbc
End Function

Public Function GetDBCTempITRBWithData() As SQLiteCConnection
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTemp()
    Dim AffectedRows As Long
    AffectedRows = CreateITRBTableWithData(dbc)
    If AffectedRows <= 0 Then Set dbc = Nothing
    Set GetDBCTempITRBWithData = dbc
End Function
