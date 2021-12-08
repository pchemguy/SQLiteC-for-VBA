Attribute VB_Name = "ADODBSQLiteNumberedParam"
'@Folder "SQLiteDBdev.Examples.Sample Parameter Query"
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


Private Type TAdoParam
    Name As String
    Value As Variant
    Size As Long
    DataType As ADODB.DataTypeEnum
End Type


'@Description "Determines ADODB Parameter Data Type for a VBA variable"
Private Function GetAdoParamType(ByVal TypeValue As Variant) As ADODB.DataTypeEnum
Attribute GetAdoParamType.VB_Description = "Determines ADODB Parameter Data Type for a VBA variable"
    Select Case VarType(TypeValue)
        Case vbString
            GetAdoParamType = adVarWChar
        Case vbInteger, vbLong
            GetAdoParamType = adInteger
        Case vbSingle, vbDouble
            GetAdoParamType = adDouble
        Case Else
            GetAdoParamType = adVarWChar
    End Select
End Function


'@Description "Selects appropriate ADODB Parameter size for a VBA variable value"
Private Function GetAdoParamSize(ByVal Param As Variant) As Long
Attribute GetAdoParamSize.VB_Description = "Selects appropriate ADODB Parameter size for a VBA variable value"
    Select Case VarType(Param)
        Case vbString
            GetAdoParamSize = Len(Param)
        Case vbInteger, vbLong
            GetAdoParamSize = 8
        Case vbSingle, vbDouble
            GetAdoParamSize = 8
        Case Else
            GetAdoParamSize = 255
    End Select
End Function


Private Function GetAdoCommand() As ADODB.Command
    Dim Database As String
    Database = ThisWorkbook.Path + "\" + "SQLiteDB.db"
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    Dim AdoConnStr As String
    AdoConnStr = "Driver=" + Driver + ";" + _
                 "Database=" + Database + ";" + _
                 Options
    Dim TableName As String
    TableName = "contacts"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & _
              " WHERE [id] <= ?2 AND [Age] < ?1 AND [Gender] = ?4 AND [Email] LIKE ?3"

    Dim FieldNames As Variant
    FieldNames = Array("Age", "id", "Email", "Gender")
    Dim FieldTypeValues As Variant
    FieldTypeValues = Array(0, 0, " ", " ")
    
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQLQuery
        .Prepared = True
        .ActiveConnection = AdoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
    
    Dim AdoParamProps As TAdoParam
    Dim AdoParam As ADODB.Parameter
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
        With AdoParamProps
            .Name = FieldNames(FieldIndex)
            .Value = FieldTypeValues(FieldIndex)
            .Size = GetAdoParamSize(.Value)
            .DataType = GetAdoParamType(.Value)
            Set AdoParam = AdoCommand.CreateParameter(.Name, .DataType, adParamInput, .Size, .Value)
            AdoCommand.Parameters.Append AdoParam
        End With
    Next FieldIndex
    
    Set GetAdoCommand = AdoCommand
End Function


'@EntryPoint
Private Sub DemoSQLite3WithPosParams()
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetAdoCommand

    Dim FieldNames As Variant
    FieldNames = Array("Age", "id", "Email", "Gender")
    Dim FieldValues As Variant
    FieldValues = Array(50, 500, "%.net", "male")
    
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
        AdoCommand.Parameters(FieldNames(FieldIndex)).Size = GetAdoParamSize(FieldValues(FieldIndex))
        AdoCommand.Parameters(FieldNames(FieldIndex)).Value = FieldValues(FieldIndex)
    Next FieldIndex
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close

    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
        
    Set WSQueryTable = Buffer.QueryTables.Add(Connection:=AdoRecordset, Destination:=Buffer.Range("A1"))
    With WSQueryTable
        .FieldNames = True
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .EnableEditing = True
    End With
    WSQueryTable.Refresh
    Buffer.UsedRange.Rows(1).HorizontalAlignment = xlCenter
End Sub


'@EntryPoint
Private Sub DemoSQLite3Ref()
    Dim Database As String
    Database = ThisWorkbook.Path + "\" + "SQLiteDB.db"
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    Dim AdoConnStr As String
    AdoConnStr = "Driver=" + Driver + ";" + _
                 "Database=" + Database + ";" + _
                 Options
    Dim TableName As String
    TableName = "contacts"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & _
              " WHERE [id] <= 500 AND [Age] < 50 AND [Gender] = 'male' AND [Email] LIKE '%.net'"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQLQuery
        .ActiveConnection = AdoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With

    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close

    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
        
    Set WSQueryTable = Buffer.QueryTables.Add(Connection:=AdoRecordset, Destination:=Buffer.Range("A1"))
    With WSQueryTable
        .FieldNames = True
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .EnableEditing = True
    End With
    WSQueryTable.Refresh
    Buffer.UsedRange.Rows(1).HorizontalAlignment = xlCenter
End Sub
