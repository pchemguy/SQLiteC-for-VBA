Attribute VB_Name = "PlainADODBSampleCode"
'@Folder "Drafts.Basic"
'@IgnoreModule
Option Explicit


Private Sub TestADODBSourceCMDCSV()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String

    Dim AdoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String

    #If Win64 Then
        sDriver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        sDriver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    sDatabaseExt = ".csv"
    sDatabase = ThisWorkbook.Path
    sTable = fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
    AdoConnStr = "Driver=" & sDriver & ";" & _
                 "DefaultDir=" & sDatabase & ";"

    sSQL = "SELECT * FROM """ & sTable & """"
    sSQL = sTable

    qtConnStr = "OLEDB;" + AdoConnStr

    sSQL = "SELECT * FROM people WHERE id <= 45 AND last_name <> 'machinery'"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .CommandText = sSQL
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
End Sub


Private Sub TestADODBSourceSQL()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim AdoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String

    sDatabase = ThisWorkbook.Path + "\" + "ADODBTemplates.db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    AdoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + AdoConnStr

    sSQL = "SELECT * FROM people WHERE id <= 45 AND last_name <> 'machinery'"

    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    On Error Resume Next
    AdoConnection.Open AdoConnStr
    On Error GoTo 0
    If AdoConnection.State = ADODB.ObjectStateEnum.adStateOpen Then AdoConnection.Close

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open Source:=sSQL, ActiveConnection:=AdoConnStr, CursorType:=adOpenKeyset, LockType:=adLockReadOnly, Options:=(adCmdText Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub TestADODBSourceCMD()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim AdoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String

    sDatabase = ThisWorkbook.Path + "\" + "ADODBTemplates.db"
    sDriver = "SQLite3 ODBC Driver"
    sOptions = "SyncPragma=NORMAL;FKSupport=True;"
    AdoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + AdoConnStr

    sSQL = "SELECT * FROM people WHERE id <= 45 AND last_name <> 'machinery'"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .CommandText = sSQL
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
End Sub


Private Sub TestADODBSourceSQLite()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim AdoConnStr As String
    Dim sSQL As String

    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & Application.PathSeparator & fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
    AdoConnStr = "Driver=" & sDriver & ";" & _
                 "Database=" & sDatabase & ";"

    sSQL = "SELECT * FROM """ & sTable & """"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=AdoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdText Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub TestADODBSourceCSV()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim AdoConnStr As String
    Dim sSQL As String

    #If Win64 Then
        sDriver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        sDriver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    sDatabaseExt = ".csv"
    sDatabase = ThisWorkbook.Path
    sTable = fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
    AdoConnStr = "Driver=" & sDriver & ";" & _
                 "DefaultDir=" & sDatabase & ";"

    sSQL = "SELECT * FROM """ & sTable & """"
    sSQL = sTable

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=AdoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdTable Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub TestADODBConnectCSV()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim AdoConnStr As String
    Dim sSQL As String

    #If Win64 Then
        sDriver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        sDriver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    sDatabaseExt = ".csv"
    sDatabase = ThisWorkbook.Path
    sTable = fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
    AdoConnStr = "Driver=" & sDriver & ";" & _
                 "DefaultDir=" & sDatabase & ";"

    sSQL = "SELECT * FROM """ & sTable & """"
    sSQL = sTable

    Dim AdoConnection As ADODB.Connection
    Set AdoConnection = New ADODB.Connection
    AdoConnection.ConnectionString = AdoConnStr

    On Error Resume Next
    AdoConnection.Open
    Debug.Print AdoConnection.Errors.Count
    Debug.Print AdoConnection.Properties("Transaction DDL")
    AdoConnection.BeginTrans
    On Error GoTo 0

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseClient
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=AdoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdTable Or adAsyncFetch)
    Set AdoRecordset.ActiveConnection = Nothing
End Sub


Private Sub TestADODBConnectSQLite()
    Dim fso As New Scripting.FileSystemObject
    Dim sDriver As String
    Dim sDatabase As String
    Dim sDatabaseExt As String
    Dim sTable As String
    Dim AdoConnStr As String
    Dim sSQL As String

    sDriver = "SQLite3 ODBC Driver"
    sDatabaseExt = ".db"
    sTable = "people"
    sDatabase = ThisWorkbook.Path & Application.PathSeparator & fso.GetBaseName(ThisWorkbook.Name) & sDatabaseExt
    AdoConnStr = "Driver=" & sDriver & ";" & _
                 "Database=" & sDatabase & ";"

    sSQL = "SELECT * FROM """ & sTable & """"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    AdoRecordset.CursorLocation = adUseServer
    AdoRecordset.Open _
            Source:=sSQL, _
            ActiveConnection:=AdoConnStr, _
            CursorType:=adOpenKeyset, _
            LockType:=adLockReadOnly, _
            Options:=(adCmdText Or adAsyncFetch)
    On Error Resume Next
    Set AdoRecordset.ActiveConnection = Nothing
    On Error GoTo 0
End Sub
