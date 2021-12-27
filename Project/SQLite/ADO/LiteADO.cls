VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Minimalistic ADODB wrapper used by SQLiteDB"
'@Folder "SQLite.ADO"
'@ModuleDescription "Minimalistic ADODB wrapper used by SQLiteDB"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, ValueRequired
Option Explicit

Private Const MODULE_NAME As String = "LiteADO"

Implements ILiteADO

Private Type TLiteADO
    MainDB As String '''' Absolute file pathname to the main database.
    ODBCOptions As Variant
    ConnectionString As String
    PragmaOptions As Scripting.Dictionary
    AdoConnection As ADODB.Connection
    AdoCommand As ADODB.Command
End Type
Private this As TLiteADO


'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Variant = False, _
              Optional ByVal ConnOptions As Variant = Empty) As ILiteADO
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteADO
    Set Instance = New LiteADO
    Instance.Init Database, AllowNonExistent, ConnOptions
    Set Create = Instance
End Function


'''' Args:
''''   Database (string):
''''     Name of the database to be opened as the "main" SQLite database.
''''
''''   AllowNonExistent (Variant, optional, False):
''''     If True, Database may refer to a non-existent database, which will be
''''     created. If Empty, path resolution will be skipped.
''''
''''   ConnOptions (Variant, optional, Empty):
''''     Either a string or dictionary of ODBC options. If a string is provided,
''''     it is used as is. If a dictionary is provided, the containing options
''''     are added to / override the default options.
''''
Friend Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Variant = False, _
       Optional ByVal ConnOptions As Variant = Empty)
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = adCmdText
        .CommandText = LiteMetaSQL.Version
        .Prepared = True
        .ActiveConnection = ConnectionString(Database, AllowNonExistent, ConnOptions)
        .ActiveConnection.CursorLocation = adUseClient
        Set this.AdoConnection = .ActiveConnection
    End With
    Set this.PragmaOptions = New Scripting.Dictionary
    this.PragmaOptions.CompareMode = TextCompare
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    this.AdoConnection.Close
    On Error GoTo 0
End Sub


'@Description "Generate ILiteADO instance from and existing ADODB.Connection."
Public Function FromConnection(ByVal db As ADODB.Connection) As ILiteADO
Attribute FromConnection.VB_Description = "Generate ILiteADO instance from and existing ADODB.Connection."
    Dim Instance As LiteADO
    Set Instance = New LiteADO
    Instance.InitFromConnection db
    Set FromConnection = Instance
End Function


'@Description "Generate ILiteADO instance from and existing ADODB.Connection."
Friend Sub InitFromConnection(ByVal db As ADODB.Connection)
Attribute InitFromConnection.VB_Description = "Generate ILiteADO instance from and existing ADODB.Connection."
    Guard.NullReference db
    
    Set this.AdoConnection = db
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = adCmdText
        .Prepared = True
        Set .ActiveConnection = db
    End With
    this.MainDB = GetScalar(SQLlib.AttachedDbPaths("main"))
    this.AdoCommand.CommandText = LiteMetaSQL.Version
End Sub


Public Property Get AdoConnection() As ADODB.Connection
    Set AdoConnection = this.AdoConnection
End Property


Public Property Get AdoCommand() As ADODB.Command
    Set AdoCommand = this.AdoCommand
End Property


Public Property Get ODBCOptions() As Variant
    If IsObject(this.ODBCOptions) Then
        Set ODBCOptions = this.ODBCOptions
    Else
        ODBCOptions = this.ODBCOptions
    End If
End Property


Public Property Get ODBCOptionsStr() As String
    Dim Options As String
    If VarType(this.ODBCOptions) = vbString Then
        Options = this.ODBCOptions
    Else
        Dim OptionsDict As Scripting.Dictionary
        Set OptionsDict = IIf(IsObject(this.ODBCOptions), this.ODBCOptions, DefaultOptionsDict)
        Dim OptionsArray As Variant
        OptionsArray = OptionsDict.Keys
        Dim OptionName As String
        Dim OptionIndex As Long
        For OptionIndex = 0 To OptionsDict.Count - 1
            OptionName = OptionsArray(OptionIndex)
            OptionsArray(OptionIndex) = OptionName & "=" & OptionsDict(OptionName)
        Next OptionIndex
        Options = Join(OptionsArray, ";") & ";"
    End If
    ODBCOptionsStr = Options
End Property


Public Property Get DefaultOptions() As String
    DefaultOptions = "StepAPI=True;NoTXN=False;NoWCHAR=False;SyncPragma=NORMAL;FKSupport=True;"
End Property


Public Property Get DefaultOptionsDict() As Scripting.Dictionary
    Dim Options As Scripting.Dictionary
    Set Options = New Scripting.Dictionary
    With Options
        .CompareMode = TextCompare
        .Item("StepAPI") = True
        .Item("NoTXN") = False
        .Item("NoWCHAR") = False
        .Item("SyncPragma") = "NORMAL"
        .Item("FKSupport") = True
    End With
    Set DefaultOptionsDict = Options
End Property


'''' Args:
''''   DatabaseName (string):
''''     File pathname of the main database
''''
''''   AllowNonExistent (boolean, optional, false):
''''     If false, creation of new database will not be allowed.
''''
''''   Options (Variant, optional, Empty):
''''     Either a string or dictionary of ODBC options. If a string is provided,
''''     it is used as is. If a dictionary is provided, the containing options
''''     are added to / override the default options.
''''
'@Description "Constructs default SQLiteODBC connection string"
Private Function ConnectionString(ByVal Database As String, _
                         Optional ByVal AllowNonExistent As Variant = False, _
                         Optional ByVal Options As Variant = Empty) As String
Attribute ConnectionString.VB_Description = "Constructs default SQLiteODBC connection string"
    this.MainDB = Database
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim OptionsStr As String
    Select Case VarType(Options)
        Case vbEmpty, vbNull
            Set this.ODBCOptions = DefaultOptionsDict
        Case vbString
            this.ODBCOptions = Options
        Case vbObject
            Dim OptionsDict As Scripting.Dictionary
            Set OptionsDict = DefaultOptionsDict
            Dim OptionNames As Variant
            OptionNames = Options.Keys()
            Dim OptionIndex As Long
            Dim OptionName As String
            For OptionIndex = 0 To Options.Count - 1
                OptionName = OptionNames(OptionIndex)
                OptionsDict(OptionName) = Options(OptionName)
            Next OptionIndex
            Set this.ODBCOptions = OptionsDict
    End Select
    If Not AllowNonExistent Then
        If IsObject(this.ODBCOptions) Then
            this.ODBCOptions("NoCreat") = True
        Else
            this.ODBCOptions = this.ODBCOptions & "NoCreat=True;"
        End If
    End If
    OptionsStr = ODBCOptionsStr
    
    this.ConnectionString = "Driver=" & Driver & ";Database=" & Database & ";" & OptionsStr
    ConnectionString = this.ConnectionString
End Function


'''' Opens, disconnects, and returns an ADODB.Recordset. If SQLQuery is provided,
'''' this.AdoCommand attribute is updated before initializing the Recordset.
''''
'''' Empty SQLQuery can be used, for example, with previously set parametrized query.
'''' Parameter values should be bound to this.AdoCommand before calling this method.
''''
'@Description "Returns disconnected Recordset"
Private Function GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
Attribute GetAdoRecordset.VB_Description = "Returns disconnected Recordset"
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    If Len(SQLQuery) > 0 Then this.AdoCommand.CommandText = SQLQuery
    With AdoRecordset
        Set .Source = this.AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
        Set .ActiveConnection = Nothing
    End With

    Set GetAdoRecordset = AdoRecordset
End Function


'@Description "Returns disconnected Recordset"
Private Function GetScalar(Optional ByVal SQLQuery As String = vbNullString) As Variant
Attribute GetScalar.VB_Description = "Returns disconnected Recordset"
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    If Len(SQLQuery) > 0 Then this.AdoCommand.CommandText = SQLQuery
    With AdoRecordset
        Set .Source = this.AdoCommand
        .MaxRecords = 1
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        Set .ActiveConnection = Nothing
    End With
    
    If AdoRecordset.RecordCount = 0 Then
        GetScalar = Empty
    Else
        GetScalar = AdoRecordset.Fields(0).Value
    End If
End Function


'@Description "Executes database modifying statement (UPDATE, DELETE, INSERT)"
Private Function ExecuteNonQuery(ByVal SQLQuery As String) As Long
Attribute ExecuteNonQuery.VB_Description = "Executes database modifying statement (UPDATE, DELETE, INSERT)"
    Dim SubQueries As Variant
    SubQueries = SQLlib.SplitMultiStatement(SQLQuery)
    Dim TotalRecordsAffected As Long
    TotalRecordsAffected = 0
    Dim RecordsAffected As Long
    Dim SubQueryIndex As Long
    For SubQueryIndex = LBound(SubQueries) To UBound(SubQueries)
        RecordsAffected = 0
        With this.AdoCommand
            .CommandText = SubQueries(SubQueryIndex)
            .Execute RecordsAffected, Options:=adExecuteNoRecords
        End With
        TotalRecordsAffected = TotalRecordsAffected + RecordsAffected
    Next SubQueryIndex
    ExecuteNonQuery = TotalRecordsAffected
End Function


'============================================================================='
'============================= ILiteADO INTERFACE ============================'
'============================================================================='


Private Property Get ILiteADO_MainDB() As String
    ILiteADO_MainDB = this.MainDB
End Property

Private Property Get ILiteADO_ConnectionString() As String
    ILiteADO_ConnectionString = this.ConnectionString
End Property

Private Property Get ILiteADO_ClassName() As String
    ILiteADO_ClassName = MODULE_NAME
End Property

Private Function ILiteADO_OpenDb() As Long
    Err.Clear
    On Error Resume Next
    If this.AdoConnection.State = adStateClosed Then this.AdoConnection.Open
    ILiteADO_OpenDb = Err.Number
    On Error GoTo 0
End Function

Private Function ILiteADO_CloseDb() As Long
    Err.Clear
    On Error Resume Next
    If this.AdoConnection.State <> adStateClosed Then this.AdoConnection.Close
    ILiteADO_CloseDb = Err.Number
    On Error GoTo 0
End Function

Private Property Get ILiteADO_ConnectionClosed() As Boolean
    If Not this.AdoConnection Is Nothing Then
        ILiteADO_ConnectionClosed = (Not this.AdoConnection.State <> adStateClosed)
    Else
        ILiteADO_ConnectionClosed = True
    End If
End Property

Private Function ILiteADO_GetScalar(ByVal SQLQuery As String, _
                           Optional ByVal QueryParameters As Variant = Empty _
                           ) As Variant
    Guard.ExpressionErr IsFalsy(QueryParameters), ActionNotSupportedErr, _
                        "LiteADO/ILiteADO_GetScalar", _
                        "Parameterized queries are not supported."
    ILiteADO_GetScalar = GetScalar(SQLQuery)
End Function

Private Function ILiteADO_GetAdoRecordset(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As ADODB.Recordset
    Guard.ExpressionErr IsFalsy(QueryParameters), ActionNotSupportedErr, _
                        "LiteADO/ILiteADO_GetAdoRecordset", _
                        "Parameterized queries are not supported."
    Set ILiteADO_GetAdoRecordset = GetAdoRecordset(SQLQuery)
End Function

Private Function ILiteADO_ExecuteNonQuery(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As Long
    Guard.ExpressionErr IsFalsy(QueryParameters), ActionNotSupportedErr, _
                        "LiteADO/ILiteADO_ExecuteNonQuery", _
                        "Parameterized queries are not supported."
    ILiteADO_ExecuteNonQuery = ExecuteNonQuery(SQLQuery)
End Function

Private Sub ILiteADO_Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED)
    If Mode = SQLITE_TXN_DEFERRED Then
        this.AdoConnection.BeginTrans
    Else
        Dim TxMode As String
        Select Case Mode
            Case SQLITE_TXN_IMMEDIATE
                TxMode = " IMMEDIATE"
            Case SQLITE_TXN_EXCLUSIVE
                TxMode = " EXCLUSIVE"
            Case Else
                Err.Raise ErrNo.InvalidParameterErr, "LiteADO", _
                    "Invalid transaction mode: <" & CStr(Mode) & ">"
        End Select
        ExecuteNonQuery "BEGIN" & TxMode
    End If
End Sub

Private Sub ILiteADO_Commit()
    this.AdoConnection.CommitTrans
End Sub

Private Sub ILiteADO_SavePoint(ByVal SavePointName As String)
    Guard.EmptyString SavePointName
    ExecuteNonQuery "SAVEPOINT " & SavePointName
End Sub

Private Sub ILiteADO_ReleasePoint(ByVal SavePointName As String)
    Guard.EmptyString SavePointName
    ExecuteNonQuery "RELEASE " & SavePointName
End Sub

Private Sub ILiteADO_Rollback(Optional ByVal SavePointName As String = vbNullString)
    If Len(SavePointName) = 0 Then
        this.AdoConnection.RollbackTrans
    Else
        ExecuteNonQuery "ROLLBACK " & SavePointName
    End If
End Sub

'@Description "Caches PragmaOptions queries."
Private Function PragmaOptionsDict(ByVal PragmaName As String) As Scripting.Dictionary
Attribute PragmaOptionsDict.VB_Description = "Caches PragmaOptions queries."
    Dim PragmaOptions As Scripting.Dictionary
    With this.PragmaOptions
        If .Exists(PragmaName) Then
            Set PragmaOptions = .Item(PragmaName)
        Else
            Set PragmaOptions = LiteMetaADO(Me).PragmaOptionsDict(PragmaName)
            Set .Item(PragmaName) = PragmaOptions
        End If
    End With
    Set PragmaOptionsDict = PragmaOptions
End Function

'@Description "Returns a dictionary with keys set to available SQLite functions."
Private Property Get ILiteADO_Functions() As Scripting.Dictionary
Attribute ILiteADO_Functions.VB_Description = "Returns a dictionary with keys set to available SQLite functions."
    Set ILiteADO_Functions = PragmaOptionsDict("function_list")
End Property

'@Description "Returns a dictionary with keys set to available SQLite pragmas."
Private Property Get ILiteADO_Pragmas() As Scripting.Dictionary
Attribute ILiteADO_Pragmas.VB_Description = "Returns a dictionary with keys set to available SQLite pragmas."
    Set ILiteADO_Pragmas = PragmaOptionsDict("pragma_list")
End Property

'@Description "Returns a dictionary with keys set to available SQLite modules."
Private Property Get ILiteADO_Modules() As Scripting.Dictionary
Attribute ILiteADO_Modules.VB_Description = "Returns a dictionary with keys set to available SQLite modules."
    Set ILiteADO_Modules = PragmaOptionsDict("module_list")
End Property

'@Description "Returns a dictionary with keys set to available SQLite compile options."
Private Property Get ILiteADO_CompileOptions() As Scripting.Dictionary
Attribute ILiteADO_CompileOptions.VB_Description = "Returns a dictionary with keys set to available SQLite compile options."
    Set ILiteADO_CompileOptions = PragmaOptionsDict("compile_options")
End Property

'@Description "Checks if a particular feature is available."
Private Function ILiteADO_FeatureAvailable(ByVal FeatureName As String) As Boolean
Attribute ILiteADO_FeatureAvailable.VB_Description = "Checks if a particular feature is available."
    Dim CompileOption As String
    Select Case UCase$(FeatureName)
        Case "JSON", "JSON1"
            CompileOption = "ENABLE_JSON1"
        Case "URI"
            CompileOption = "USE_URI"
        Case Else
            CompileOption = "ENABLE_" & FeatureName
    End Select
    Dim Status As Boolean
    '@Ignore AssignmentNotUsed: False positive
    Status = False
    With PragmaOptionsDict("compile_options")
        If .Exists(CompileOption) Then Status = CBool(.Item(CompileOption))
    End With
    ILiteADO_FeatureAvailable = Status
End Function
