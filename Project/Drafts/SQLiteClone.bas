Attribute VB_Name = "SQLiteClone"
'@Folder "Drafts"
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, AssignmentNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, SelfAssignedDeclaration
Option Explicit


Private Function GetDatabasePath(Optional ByVal DatabaseName As String = vbNullString) As String
    Dim DatabasePath As String
    If DatabaseName = vbNullString Then
        DatabasePath = ThisWorkbook.Path & Application.PathSeparator & "ADODBTemplates.db"
    ElseIf DatabaseName = "cache" Then
        DatabasePath = ":memory:"
    End If
    GetDatabasePath = DatabasePath
End Function


Private Function GetConnectionString(Optional ByVal DatabaseName As String = vbNullString) As Scripting.Dictionary
    Dim Driver As String: Driver = "SQLite3 ODBC Driver"

    Dim Database As String
    Database = GetDatabasePath(DatabaseName)
    
    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    ConnectionString("OLEDB") = "OLEDB;" + ConnectionString("ADO")

    Set GetConnectionString = ConnectionString
End Function


Private Function GetSQL() As Scripting.Dictionary
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Dim fso As New Scripting.FileSystemObject
    SQL("TableName") = "people"
    SQL("Query") = "SELECT * FROM [" & SQL("TableName") & "] WHERE id >= 10 AND id < 20"
    SQL("Query") = "SELECT id, FirstName As FirstName, LastName As LastName FROM [" & SQL("TableName") & "]" ' WHERE id >= 10 AND id < 20"
    Set GetSQL = SQL
End Function


Private Function GetAdoCommand(Optional ByVal DatabaseName As String = vbNullString) As ADODB.Command
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString(DatabaseName)

    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQL

    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQL("Query")
        .ActiveConnection = ConnectionString("ADO")
        .ActiveConnection.CursorLocation = adUseClient
    End With
    Set GetAdoCommand = AdoCommand
End Function


Private Function GetAdoRecordset(Optional ByVal DatabaseName As String = vbNullString) As ADODB.Recordset
    Dim AdoCommand As ADODB.Command: Set AdoCommand = GetAdoCommand(DatabaseName)
    Dim AdoRecordset As ADODB.Recordset: Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = GetAdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With

    Set GetAdoRecordset = AdoRecordset
End Function


Private Sub RecordSetSourceAdoCommand()
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset(vbNullString)
End Sub


Private Sub InMemoryDb()
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetAdoCommand("cache")
End Sub


Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = vbNullString
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetAdoCommand(SourceDb)

    Dim Catalog As ADOX.Catalog
    Set Catalog = New ADOX.Catalog
    Set Catalog.ActiveConnection = AdoCommand.ActiveConnection
    
End Sub
