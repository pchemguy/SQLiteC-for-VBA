Attribute VB_Name = "SQLiteUpdateRecordset"
'@Folder "SQLiteDB.Drafts"
'@IgnoreModule VariableNotUsed, AssignmentNotUsed, ParameterNotUsed, ProcedureNotUsed, SelfAssignedDeclaration
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
Option Explicit

Private Type TSQLiteUpdateRecordset
    ConnectionString As String
    TableName As String
    TableMetaLoaded As Boolean
    AdoCommand As ADODB.Command
    FieldNames() As String
    FieldTypes() As ADODB.DataTypeEnum
End Type
Private this As TSQLiteUpdateRecordset


Private Function GetConnectionString() As String
    Dim Driver As String: Driver = "SQLite3 ODBC Driver"

    Dim Database As String
    Database = ThisWorkbook.Path & Application.PathSeparator & "ADODBTemplates.db"

    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    
    GetConnectionString = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
End Function


Private Sub Init()
    this.ConnectionString = GetConnectionString
    this.TableName = "people"
    this.TableMetaLoaded = False
End Sub


Public Function SQLSelectWildcard(Optional ByVal SelectLimit As Long = 0) As String
    If this.TableName = vbNullString Then Init
    Dim LimitClause As String
    LimitClause = IIf(SelectLimit > 0, " LIMIT " & CStr(SelectLimit), vbNullString)
    SQLSelectWildcard = "SELECT * FROM [" & this.TableName & "]" & LimitClause
End Function


Public Function SQLSelectAllAliased(Optional ByVal SelectLimit As Long = 0) As String
    If Not this.TableMetaLoaded Then GetTableMeta
    
    Dim LimitClause As String
    LimitClause = IIf(SelectLimit > 0, " LIMIT " & CStr(SelectLimit), vbNullString)
    
    Dim FieldCount As Long
    FieldCount = UBound(this.FieldNames)
    Dim FieldAliases() As String
    ReDim FieldAliases(1 To FieldCount)
    
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        FieldAliases(FieldIndex) = "[" & this.FieldNames(FieldIndex) & "] AS [" & this.FieldNames(FieldIndex) & "]"
    Next FieldIndex
    SQLSelectAllAliased = "SELECT " & Join(FieldAliases, ", ") & " FROM [" & this.TableName & "]" & LimitClause
End Function


Public Function SQLSelectAllAliasedAsText(Optional ByVal SelectLimit As Long = 0) As String
    If Not this.TableMetaLoaded Then GetTableMeta
    
    Dim LimitClause As String
    LimitClause = IIf(SelectLimit > 0, " LIMIT " & CStr(SelectLimit), vbNullString)
    
    Dim FieldCount As Long
    FieldCount = UBound(this.FieldNames)
    Dim FieldAliases() As String
    ReDim FieldAliases(1 To FieldCount)
    
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        If this.FieldTypes(FieldIndex) = adVarWChar Then
            FieldAliases(FieldIndex) = "[" & this.FieldNames(FieldIndex) & "] AS [" & this.FieldNames(FieldIndex) & "]"
        Else
            FieldAliases(FieldIndex) = "CAST([" & this.FieldNames(FieldIndex) & "] AS TEXT) AS [" & this.FieldNames(FieldIndex) & "]"
        End If
    Next FieldIndex
    SQLSelectAllAliasedAsText = "SELECT " & Join(FieldAliases, ", ") & " FROM [" & this.TableName & "]" & LimitClause
End Function


Private Function GetAdoCommand(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Command
    If this.TableName = vbNullString Then Init
    
    On Error Resume Next
    this.AdoCommand.ActiveConnection.Close
    On Error GoTo 0
    
    Dim Query As String
    Query = IIf(Len(SQLQuery) > 0, SQLQuery, SQLSelectWildcard)

    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = Query
        .ActiveConnection = this.ConnectionString
        .ActiveConnection.CursorLocation = adUseClient
    End With
    Set this.AdoCommand = AdoCommand
    Set GetAdoCommand = AdoCommand
End Function


Private Function GetDisconnectedAdoRecordset(Optional ByVal SQLQuery As String = vbNullString, _
                                             Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetAdoCommand(SQLQuery)
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
        .MarshalOptions = adMarshalModifiedOnly
        Set .ActiveConnection = Nothing
    End With

    Set GetDisconnectedAdoRecordset = AdoRecordset
End Function


Private Sub GetTableMeta()
    If this.ConnectionString = vbNullString Then Init

    Dim Catalog As New ADOX.Catalog
    Catalog.ActiveConnection = this.ConnectionString
    
    Dim AdoTable As ADOX.Table
    Set AdoTable = Catalog.Tables(this.TableName)
    Dim ColumnCount As Long
    ColumnCount = AdoTable.Columns.Count
    
    ReDim this.FieldNames(1 To ColumnCount)
    ReDim this.FieldTypes(1 To ColumnCount)
    
    Dim Column As ADOX.Column
    Dim ColumnIndex As Long
    For ColumnIndex = 0 To ColumnCount - 1
        Set Column = AdoTable.Columns(ColumnIndex)
        this.FieldNames(ColumnIndex + 1) = Column.Name
        this.FieldTypes(ColumnIndex + 1) = Column.Type
    Next ColumnIndex
    
    this.TableMetaLoaded = True
End Sub


Private Sub UpdateRowsBatch()
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetDisconnectedAdoRecordset(SQLSelectAllAliased, adLockBatchOptimistic)
        
    Dim FirstName As String
    With AdoRecordset
        Debug.Print .Supports(adAddNew)
        .MoveFirst
        
        .Find "id = 15"
        FirstName = .Fields("FirstName")
        '@Ignore SuspiciousLetAssignment
        .Fields("FirstName") = .Fields("LastName")
        .Fields("LastName") = FirstName
        FirstName = .Fields("FirstName")
        
        .Find "id = 20"
        FirstName = .Fields("FirstName")
        '@Ignore SuspiciousLetAssignment
        .Fields("FirstName") = .Fields("LastName")
        .Fields("LastName") = FirstName
        FirstName = .Fields("FirstName")
        
        .Find "id > 1000"
        Do While Not .EOF
            '@Ignore ValueRequired
            Debug.Print .AbsolutePosition, .Fields("id")
            .Delete
            .MoveNext
        Loop
        Set .ActiveConnection = this.AdoCommand.ActiveConnection
        .UpdateBatch
        Set .ActiveConnection = Nothing
        
        Dim Suffix As Long
        Suffix = 1
        RstAddNew AdoRecordset, 3
        RstAddNew AdoRecordset, 4
        RstAddNew AdoRecordset, 5
        
        Set .ActiveConnection = this.AdoCommand.ActiveConnection
        .UpdateBatch
    End With
End Sub


Private Sub RstAddNew(ByVal AdoRecordset As ADODB.Recordset, ByVal Suffix As Long)
        With AdoRecordset
            .AddNew
            .Fields("FirstName") = "FirstName" & Suffix
            .Fields("LastName") = "LastName" & Suffix
            .Fields("Age") = "64"
            .Fields("Gender") = "male"
            .Fields("Email") = "FirstName" & Suffix & ".LastName" & Suffix & "@domain" & Suffix & ".com"
            .Fields("Country") = "Country" & Suffix
            .Fields("Domain") = "domain" & Suffix & ".com"
        End With
End Sub
