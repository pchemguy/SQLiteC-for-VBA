Attribute VB_Name = "CSVTemplate"
'@Folder "Drafts.Basic"
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration, AssignmentNotUsed
Option Explicit


Private Function GetConnectionString() As Scripting.Dictionary
    Dim Driver As String
    #If Win64 Then
        Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If

    Dim Database As String
    Database = ThisWorkbook.Path

    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "DefaultDir=" & Database & ";"
    ConnectionString("OLEDB") = "OLEDB;" + ConnectionString("ADO")
    ConnectionString("TEXT") = "TEXT;" & Database & Application.PathSeparator

    Set GetConnectionString = ConnectionString
End Function


Private Function GetSQL() As Scripting.Dictionary
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Dim DatabaseExt As String: DatabaseExt = ".txt"
    Dim fso As New Scripting.FileSystemObject
    SQL("TableName") = fso.GetBaseName(ThisWorkbook.Name) & DatabaseExt
    SQL("Query") = "SELECT * FROM """ & SQL("TableName") & """ WHERE id <= 40"
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


Private Sub QueryTableSourceSQLCommandTableName()
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQL

    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString

    Buffer.Range(SQL("TableName")).CurrentRegion.ClearContents

    Dim WSheetTable As Excel.QueryTable
    Set WSheetTable = Buffer.QueryTables.Add(Connection:=ConnectionString("OLEDB"), Destination:=Buffer.Range(SQL("TableName")))
    With WSheetTable
        .CommandType = xlCmdTable
        .CommandText = SQL("TableName")
        .Name = SQL("TableName")
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
        .Refresh BackgroundQuery:=False
    End With
End Sub


Private Sub QueryTableSourceSQLCommandSQL()
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQL

    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString

    Buffer.Range(SQL("TableName")).CurrentRegion.ClearContents

    Dim WSheetTable As Excel.QueryTable
    Set WSheetTable = Buffer.QueryTables.Add(Connection:=ConnectionString("OLEDB"), Destination:=Buffer.Range(SQL("TableName")))
    With WSheetTable
        .CommandType = xlCmdSql
        .CommandText = SQL("Query")
        .Name = SQL("TableName")
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
        .Refresh BackgroundQuery:=False
    End With
End Sub


Private Sub QueryTableSourceFile()
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQL

    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString

    Buffer.Range(SQL("TableName")).CurrentRegion.ClearContents

    Dim WSheetTable As Excel.QueryTable
    Set WSheetTable = Buffer.QueryTables.Add(Connection:=ConnectionString("TEXT") & SQL("TableName"), Destination:=Buffer.Range(SQL("TableName")))
    With WSheetTable
        .Name = SQL("TableName")
        .FieldNames = True
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFileCommaDelimiter = True
        .TextFileOtherDelimiter = ","
    End With
    WSheetTable.Refresh
End Sub


Private Sub QueryTableSourceAdoRecordset()
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQL

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset

    Buffer.Range(SQL("TableName")).CurrentRegion.ClearContents

    Dim WSheetTable As Excel.QueryTable
    Set WSheetTable = Buffer.QueryTables.Add(Connection:=AdoRecordset, Destination:=Buffer.Range(SQL("TableName")))
    With WSheetTable
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
    WSheetTable.Refresh
End Sub
