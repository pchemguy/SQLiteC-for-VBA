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


Public Function GetDefaultDBM() As SQLiteC
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim dbm As SQLiteC
    '''' Using default library names hardcoded in the SQLiteC constructor.
    Set dbm = SQLiteC(DllPath)
    If dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteC instance."
    Set GetDefaultDBM = dbm
End Function

Public Function GetConnection(ByVal DbPathName As String) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = GetDefaultDBM()
    '''' Keep a reference to the DllManager to prevent library unloading due to object destruction
    Set this.DllMan = dbm.DllMan
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteCConnection instance."
    Set GetConnection = DbConn
End Function

Public Function GetConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set GetConnDbRegular = GetConnection(DbPathName)
End Function

Public Function GetConnDbReadOnlyAttr() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & "ReadOnlyAttr.db"
    Set GetConnDbReadOnlyAttr = GetConnection(DbPathName)
End Function

Public Function GetConnDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetConnDbMemory = GetConnection(DbPathName)
End Function

Public Function GetConnDbTemp() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set GetConnDbTemp = GetConnection(DbPathName)
End Function

Public Function GetConnDbInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set GetConnDbInvalidPath = GetConnection(DbPathName)
End Function

Public Function GetConnDbMemoryWithTable() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetConnDbMemoryWithTable = GetConnection(DbPathName)
    GetConnDbMemoryWithTable.ExecuteNonQueryPlain SQLiteCTestFixSQL.CREATETableITRB
End Function

Public Function CreateFunctionsTableWithData(ByVal dbc As SQLiteCConnection) As Long
    Dim SQLQuery As String
    SQLQuery = SQLiteCTestFixSQL.CreateFunctionsTableWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    CreateFunctionsTableWithData = AffectedRows
End Function

Public Function ByteArray(ParamArray Values() As Variant) As Byte()
    Dim Buffer() As Byte
    ReDim Buffer(0 To UBound(Values))
    Dim ByteIndex As Long
    For ByteIndex = 0 To UBound(Values)
        Buffer(ByteIndex) = CByte(Values(ByteIndex))
    Next ByteIndex
    ByteArray = Buffer
End Function
