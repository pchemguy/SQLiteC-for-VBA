Attribute VB_Name = "Performance"
'@Folder "SQLite.Performance"
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
Option Explicit

Private Const CYCLE_COUNT As Long = 10 ^ 2
Private ROW_COUNT As Long


'@EntryPoint
Private Sub RunGetScalar()
    ADODBPlainScalarSQLite
    SQLiteCScalarSQLite
End Sub


'@EntryPoint
Private Sub RunGetRecordset()
    ROW_COUNT = 20
    ADODBPlain2DSQLite
    SQLiteC2DSQLite
End Sub


Private Sub ADODBPlainScalarSQLite()
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim ODBCOptions As String
    ODBCOptions = "SyncPragma=NORMAL;FKSupport=True;"
    Dim DbName As String
    DbName = ":memory:"
    Dim AdoConnStr As String
    AdoConnStr = "Driver=" + Driver + ";" + _
                 "Database=" + DbName + ";" + _
                 ODBCOptions
    Dim SQLQuery As String
    SQLQuery = "SELECT 1024"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQLQuery
        .Prepared = True
        .ActiveConnection = AdoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With

    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
    End With
    
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    With AdoRecordset
        For CycleIndex = 0 To CYCLE_COUNT
            .Requery
        Next CycleIndex
    End With
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    Dim Result As Variant
    Result = AdoRecordset.Fields(0)
    
    Debug.Print "Plain ADODB GetScalar: " & Result & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
    
    AdoRecordset.Close
    Set AdoRecordset.Source = Nothing
    Set AdoRecordset.ActiveConnection = Nothing
    AdoCommand.ActiveConnection.Close
End Sub


Private Sub SQLiteCScalarSQLite()
    Dim DbName As String
    DbName = ":memory:"
    Dim SQLQuery As String
    SQLQuery = "SELECT 1024"

    Dim dbm As SQLiteC
    Set dbm = SQLiteC(vbNullString)
    Dim dbc As SQLiteCConnection
    Set dbc = SQLiteCConnection(DbName)
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    dbc.OpenDb
    
    Dim Result As Variant
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    With dbs
        For CycleIndex = 0 To CYCLE_COUNT
            Result = .GetScalar(SQLQuery)
        Next CycleIndex
    End With
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    dbc.CloseDb
    
    Debug.Print "SQLiteC     GetScalar: " & Result & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub


Private Sub ADODBPlain2DSQLite()
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim ODBCOptions As String
    ODBCOptions = "SyncPragma=NORMAL;FKSupport=True;"
    Dim DbName As String
    DbName = ":memory:"
    Dim AdoConnStr As String
    AdoConnStr = "Driver=" + Driver + ";" + _
                 "Database=" + DbName + ";" + _
                 ODBCOptions
    Dim SQLQuery As String

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .Prepared = True
        .ActiveConnection = AdoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    
        SQLQuery = Join(Array( _
            "CREATE TABLE functions(", _
            "    name    TEXT COLLATE NOCASE NOT NULL,", _
            "    builtin INTEGER             NOT NULL,", _
            "    type    TEXT COLLATE NOCASE NOT NULL,", _
            "    enc     TEXT COLLATE NOCASE NOT NULL,", _
            "    narg    INTEGER             NOT NULL,", _
            "    flags   INTEGER             NOT NULL", _
            ");" _
        ), vbNewLine)
        .CommandText = SQLQuery
        .Execute
        SQLQuery = Join(Array( _
            "INSERT INTO functions ", _
            "SELECT * FROM pragma_function_list() ORDER BY name, flags" _
        ), vbNewLine)
        .CommandText = SQLQuery
        .Execute
        SQLQuery = "SELECT * FROM functions LIMIT " & CStr(ROW_COUNT)
        .CommandText = SQLQuery
    End With

    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
    End With
    
    Dim Result As Variant
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    With AdoRecordset
        For CycleIndex = 0 To CYCLE_COUNT
            .Requery
            Result = AdoRecordset.GetRows
        Next CycleIndex
    End With
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    
    Debug.Print "Plain ADODB GetRecordset: " & UBound(Result, 2) + 1 & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
    
    AdoRecordset.Close
    Set AdoRecordset.Source = Nothing
    Set AdoRecordset.ActiveConnection = Nothing
    AdoCommand.ActiveConnection.Close
End Sub


Private Sub SQLiteC2DSQLite()
    Dim DbName As String
    DbName = ":memory:"
    Dim SQLQuery As String

    Dim dbm As SQLiteC
    Set dbm = SQLiteC(vbNullString)
    Dim dbc As SQLiteCConnection
    Set dbc = SQLiteCConnection(DbName)
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    dbc.OpenDb
    
    SQLQuery = Join(Array( _
        "CREATE TABLE functions(", _
        "    name    TEXT COLLATE NOCASE NOT NULL,", _
        "    builtin INTEGER             NOT NULL,", _
        "    type    TEXT COLLATE NOCASE NOT NULL,", _
        "    enc     TEXT COLLATE NOCASE NOT NULL,", _
        "    narg    INTEGER             NOT NULL,", _
        "    flags   INTEGER             NOT NULL", _
        ");" _
    ), vbNewLine) & _
    Join(Array( _
        "INSERT INTO functions ", _
        "SELECT * FROM pragma_function_list() ORDER BY name, flags" _
    ), vbNewLine)
    dbc.ExecuteNonQueryPlain SQLQuery
    SQLQuery = "SELECT * FROM functions LIMIT " & CStr(ROW_COUNT)
    
    Dim Result As Variant
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    With dbs
        For CycleIndex = 0 To CYCLE_COUNT
            Result = .GetRowSet2D(SQLQuery)
        Next CycleIndex
    End With
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    dbc.CloseDb
    
    Debug.Print "SQLiteC     GetRecordset: " & UBound(Result, 1) + 1 & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub


