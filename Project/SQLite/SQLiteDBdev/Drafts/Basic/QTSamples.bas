Attribute VB_Name = "QTSamples"
'@Folder "SQLite.SQLiteDBdev.Drafts.Basic"
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration, AssignmentNotUsed, ImplicitDefaultMemberAccess
Option Explicit


Private Function GetConnectionString() As String
    Dim Driver As String
    Dim Options As String
    Dim Database As String
    
    Database = ThisWorkbook.Path + "\" + "SQLiteDB.db"
    Driver = "SQLite3 ODBC Driver"
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    
    GetConnectionString = "Driver=" + Driver + ";" + "Database=" + Database + ";" + Options
End Function


Private Function GetSQLSelectAliased() As String
    Dim TableName As String
    TableName = "people"

    Dim FieldNames() As String
    
    Dim Catalog As ADOX.Catalog
    Set Catalog = New ADOX.Catalog
    Catalog.ActiveConnection = GetConnectionString
    
    Dim Table As ADOX.Table
    Set Table = Catalog.Tables(TableName)
    ReDim FieldNames(0 To Table.Columns.Count - 1)
    
    Dim ColumnIndex As Long
    Dim FieldName As String
    For ColumnIndex = 0 To Table.Columns.Count - 1
        FieldName = Table.Columns(ColumnIndex).Name
        FieldNames(ColumnIndex) = "[" & FieldName & "] AS [" & FieldName & "]"
    Next ColumnIndex
    
    GetSQLSelectAliased = "SELECT " & Join(FieldNames, ", ") & " " & _
             "FROM [" & TableName & "] " & _
             "WHERE id <= 2000 " & _
             "ORDER BY [id] DESC" ' & "ORDER BY [Gender] DESC, [LastName] ASC, [FirstName] ASC"
End Function


Private Function GetSQLSelect() As String
    Dim TableName As String
    TableName = "people"

    GetSQLSelect = "SELECT * " & _
                   "FROM [" & TableName & "] " & _
                   "WHERE id <= 2000"
End Function


Private Sub QueryTableSourceAdoRecordset()
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    
    With AdoRecordset
        .ActiveConnection = GetConnectionString
        .Source = GetSQLSelectAliased
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
        Set .ActiveConnection = Nothing
    End With
    
    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
    
    AdoRecordset.Find "[id] > 1000"
    AdoRecordset.Fields("id") = AdoRecordset.Fields("id") + 1
    AdoRecordset.Delete
    
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


Private Sub QueryTableSourceConnStr()
    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
    
    Set WSQueryTable = Buffer.QueryTables.Add( _
                                              Connection:="OLEDB;" & GetConnectionString, _
                                              Destination:=Buffer.Range("A1"), _
                                              SQL:=GetSQLSelectAliased _
                                              )
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
End Sub


