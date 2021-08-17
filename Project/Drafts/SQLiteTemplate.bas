Attribute VB_Name = "SQLiteTemplate"
'@Folder "Drafts"
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, AssignmentNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, SelfAssignedDeclaration
Option Explicit


Private Function GetConnectionString() As Scripting.Dictionary
    Dim Driver As String: Driver = "SQLite3 ODBC Driver"

    Dim Database As String
    Database = ThisWorkbook.Path & Application.PathSeparator & "ADODBTemplates.db"

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
    SQL("Query") = "SELECT id, FirstName As FirstName, LastName As LastName FROM [" & SQL("TableName") & "] WHERE id >= 10 AND id < 20"
    Set GetSQL = SQL
End Function


Private Function GetAdoCommand() As ADODB.Command
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString

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


Private Function GetAdoRecordset() As ADODB.Recordset
    Dim AdoCommand As ADODB.Command: Set AdoCommand = GetAdoCommand
    Dim AdoRecordset As ADODB.Recordset: Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = GetAdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close

    Set GetAdoRecordset = AdoRecordset
End Function


Private Sub RecordSetSourceAdoCommand()
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset
End Sub


Private Sub ModifyAdoRecordset()
    Dim AdoCommand As ADODB.Command: Set AdoCommand = GetAdoCommand
    Dim AdoRecordset As ADODB.Recordset: Set AdoRecordset = New ADODB.Recordset
    
    With AdoRecordset
        Set .Source = AdoCommand
'        .ActiveConnection = GetConnectionString("ADO")
'        .source = GetSQL("Query")
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        Debug.Print .Supports(adIndex)
        Debug.Print .Supports(adSeek)
        .Open
        .ActiveConnection = Nothing
    End With

    With AdoRecordset
        .MoveFirst
        Debug.Print .Fields("id")
        .Find "id = 15"
        Dim FirstName As String
        FirstName = .Fields("FirstName")
        '@Ignore SuspiciousLetAssignment
        .Fields("FirstName") = .Fields("LastName")
        .Fields("LastName") = FirstName
        FirstName = .Fields("FirstName")
        .UpdateBatch
    End With
    
'    AdoRecordset.MoveFirst
'    Debug.Print AdoRecordset.Fields("people.id")
'    AdoRecordset.Find "people.id = 15"
'    FirstName = AdoRecordset.Fields("people.FirstName")
'    AdoRecordset.Fields("people.FirstName") = AdoRecordset.Fields("people.LastName")
'    AdoRecordset.Fields("people.LastName") = FirstName
'    FirstName = AdoRecordset.Fields("people.FirstName")

    With AdoRecordset
        Set .ActiveConnection = AdoCommand.ActiveConnection
        .MarshalOptions = adMarshalAll
        .Update
        .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close
End Sub
