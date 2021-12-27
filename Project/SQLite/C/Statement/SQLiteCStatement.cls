VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Manages the lifecycle of the SQLite3 Prerpared Statement Object (PSO)."
'@Folder "SQLite.C.Statement"
'@ModuleDescription "Manages the lifecycle of the SQLite3 Prerpared Statement Object (PSO)."
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Const MODULE_NAME As String = "SQLiteCStatement"

Implements ILiteADO

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, _
    ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_sql Lib "SQLite3" (ByVal hStmt As LongPtr) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_expanded_sql Lib "SQLite3" (ByVal hStmt As LongPtr) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_normalized_sql Lib "SQLite3" (ByVal hStmt As LongPtr) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_next_stmt Lib "SQLite3" (ByVal hDb As LongPtr, ByVal hStmt As LongPtr) As LongPtr ' PtrStmt
Private Declare PtrSafe Function sqlite3_stmt_busy Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_stmt_readonly Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As LongPtr)
#Else
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3" (ByVal hDb As Long, ByVal pwsSql As Long, _
    ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As SQLiteResultCodes
Private Declare Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Function sqlite3_expanded_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Function sqlite3_normalized_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Function sqlite3_next_stmt Lib "SQLite3" (ByVal hDb As Long, ByVal hStmt As Long) As Long ' PtrStmt
Private Declare Function sqlite3_stmt_busy Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_stmt_readonly Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As Long)
#End If

Private Type TSQLiteCStatement
    DbConn As SQLiteCConnection
    DbExec As SQLiteCExecSQL
    DbPara As SQLiteCParameters
    #If VBA7 Then
        StmtHandle As LongPtr
    #Else
        StmtHandle As Long
    #End If
    PragmaOptions As Scripting.Dictionary
    AffectedRowsCount As Long
    ParameterCount As Long
    PrepareOnly As Boolean
End Type
Private this As TSQLiteCStatement


'@DefaultMember
Public Function Create(ByVal DbConn As SQLiteCConnection) As SQLiteCStatement
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCStatement
    Set Instance = New SQLiteCStatement
    Instance.Init DbConn
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbConn As SQLiteCConnection)
    Guard.NullReference DbConn
    Set this.DbConn = DbConn
    Set this.DbExec = SQLiteCExecSQL(Me)
    Set this.PragmaOptions = New Scripting.Dictionary
    this.PragmaOptions.CompareMode = TextCompare
End Sub

Friend Sub CleanUp()
    With this
        '@Ignore FunctionReturnValueDiscarded
        If .StmtHandle <> 0 Then Finalize
        If Not .DbPara Is Nothing Then .DbPara.CleanUp
        If Not .DbExec Is Nothing Then .DbExec.CleanUp
        Set .DbPara = Nothing
        Set .DbExec = Nothing
        Set .DbConn = Nothing
    End With
End Sub

Friend Property Get DbConnection() As SQLiteCConnection
    Set DbConnection = this.DbConn
End Property

Friend Property Get DbExecutor() As SQLiteCExecSQL
    Set DbExecutor = this.DbExec
End Property

Friend Property Get DbParameters() As SQLiteCParameters
    Set DbParameters = this.DbPara
End Property

Public Property Get ExecADO() As ILiteADO
    Set ExecADO = Me
End Property

#If VBA7 Then
Public Property Get StmtHandle() As LongPtr
#Else
Public Property Get StmtHandle() As Long
#End If
    StmtHandle = this.StmtHandle
End Property

Public Property Get AffectedRowsCount() As Long
    AffectedRowsCount = this.AffectedRowsCount
End Property

'''' https://www.sqlite.org/c3ref/stmt_busy.html
Public Property Get Busy() As Boolean
    If this.StmtHandle <> 0 Then
        Busy = sqlite3_stmt_busy(this.StmtHandle)
    Else
        Busy = False
    End If
End Property

'''' https://www.sqlite.org/c3ref/stmt_readonly.html
Public Property Get ReadOnly() As Boolean
    If this.StmtHandle <> 0 Then
        ReadOnly = sqlite3_stmt_readonly(this.StmtHandle)
    Else
        ReadOnly = False
    End If
End Property

'''' https://sqlite.org/pragma.html#pragma_application_id
Public Property Get ApplicationId() As Long
    If this.StmtHandle <> 0 Or this.DbConn.DbHandle = 0 Then
        ApplicationId = 0
        Exit Property
    Else
        ApplicationId = GetScalar("PRAGMA application_id")
    End If
End Property

'''' https://sqlite.org/pragma.html#pragma_application_id
Public Property Let ApplicationId(ByVal Value As Long)
    With this.DbConn
        If .DbHandle <> 0 Then
            .ExecuteNonQueryPlain "PRAGMA application_id=" & CStr(Value)
        End If
    End With
End Property

'''' https://sqlite.org/pragma.html#pragma_user_version
Public Property Get UserVersion() As Long
    If this.StmtHandle <> 0 Or this.DbConn.DbHandle = 0 Then
        UserVersion = 0
        Exit Property
    Else
        UserVersion = GetScalar("PRAGMA user_version")
    End If
End Property

'''' https://sqlite.org/pragma.html#pragma_user_version
Public Property Let UserVersion(ByVal Value As Long)
    With this.DbConn
        If .DbHandle <> 0 Then
            .ExecuteNonQueryPlain "PRAGMA user_version=" & CStr(Value)
        End If
    End With
End Property


'''' Reference: https://www.sqlite.org/c3ref/prepare.html
''''
'@Ignore UseMeaningfulName
'@Description "Generates SQLite PSO."
Friend Function Prepare16V2(ByVal SQLQuery As String) As SQLiteResultCodes
Attribute Prepare16V2.VB_Description = "Generates SQLite PSO."
    FixGuard.DbNotOpened this.DbConn, MODULE_NAME & "/Prepare16V2"
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_prepare16_v2(this.DbConn.DbHandle, StrPtr(SQLQuery), _
                                      Len(SQLQuery) * 2 + 2, this.StmtHandle, 0)
    If ResultCode <> SQLITE_OK Then
        this.StmtHandle = 0
        ResultCode = this.DbConn.ErrInfoRetrieve
    Else
        If this.DbPara Is Nothing Then
            Set this.DbPara = SQLiteCParameters(Me)
            If this.DbPara Is Nothing Then
                Err.Raise ErrNo.InvalidParameterErr, MODULE_NAME & "/Prepare16V2", _
                          "Failed to instantiate SQLiteCParameters."
            End If
        End If
        '''' Retrieve parameter metadata
        this.ParameterCount = this.DbPara.ParameterCount
        If this.ParameterCount > 0 Then this.DbPara.MapParams
    End If
    Prepare16V2 = ResultCode
End Function


'''' Reference: https://www.sqlite.org/c3ref/reset.html
''''
'''' Reset should be called after execution before the statement object can be
'''' reused (for example, with a new set of bound parameters).
''''
'@Description "Resets SQLite PSO after execution."
Friend Function Reset() As SQLiteResultCodes
Attribute Reset.VB_Description = "Resets SQLite PSO after execution."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_reset(this.StmtHandle)
    If ResultCode <> SQLITE_OK Then ResultCode = this.DbConn.ErrInfoRetrieve
    Reset = ResultCode
End Function


'''' Reference: https://www.sqlite.org/c3ref/finalize.html
''''
'''' sqlite3_finalize should only return an error if the the most recent
'''' evaluation of the statement failed. At the same time, if Prepare16V2
'''' returns an error, StmtHandle is set to 0 to prevent an error here. The
'''' only remaining documented possibility is an error generated in course of
'''' automatic recompilation (https://www.sqlite.org/c3ref/prepare.html).
''''
'@Description "Destructs SQLite PSO."
Friend Function Finalize() As SQLiteResultCodes
Attribute Finalize.VB_Description = "Destructs SQLite PSO."
    If Not this.DbPara Is Nothing Then this.DbPara.Finalize
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_finalize(this.StmtHandle)
    If ResultCode = SQLITE_OK Then
        this.StmtHandle = 0
        Logger.Logg "Statement is finalized", , DEBUGLEVEL_INFO
    Else
        ResultCode = this.DbConn.ErrInfoRetrieve
    End If
    Finalize = ResultCode
End Function


'''' Reference: https://www.sqlite.org/c3ref/stmt_busy.html
'''' Reference: https://www.sqlite.org/c3ref/next_stmt.html
''''
'''' This method should be called on the predeclared SQLiteCStatement instance.
'''' It enumerates not finalized statements and either attempts to finalize
'''' them, if requested. Otherwise, it queries the busy status and returns a
'''' Dictionary mapping <statement handle> -> <busy status>.
''''
'''' Args:
''''   DbConn (SQLiteCConnection):
''''     Connection object
''''
''''   FinStmt (optional, boolean, true):
''''     If true, finalize statements
''''
'''' @ClassMethod "Call on the predeclared instance"
'@Description "Enumerates prepared statements with busy status."
Friend Function PreparedStatementsGet(ByVal DbConn As SQLiteCConnection, _
                             Optional ByVal FinStmt As Boolean = True) As Variant
Attribute PreparedStatementsGet.VB_Description = "Enumerates prepared statements with busy status."
    If DbConn.DbHandle = 0 Then
        PreparedStatementsGet = SQLITE_OK
        Exit Function
    End If
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = 0
    hStmt = sqlite3_next_stmt(DbConn.DbHandle, hStmt)
    If hStmt = 0 Then
        If FinStmt Then
            PreparedStatementsGet = SQLITE_OK
        Else
            Set PreparedStatementsGet = Nothing
        End If
        Exit Function
    End If
    
    Dim PrepStmts As Scripting.Dictionary
    Set PrepStmts = New Scripting.Dictionary
    PrepStmts.CompareMode = TextCompare
    Dim ResultCode As SQLiteResultCodes
    
    Do While hStmt <> 0
        If FinStmt Then
            ResultCode = sqlite3_finalize(hStmt)
            If ResultCode = SQLITE_OK Then
                Logger.Logg "Statement is finalized", , DEBUGLEVEL_INFO
            Else
                ResultCode = DbConn.ErrInfoRetrieve
            End If
        Else
            PrepStmts(hStmt) = sqlite3_stmt_busy(hStmt)
        End If
        hStmt = sqlite3_next_stmt(DbConn.DbHandle, hStmt)
    Loop
    If FinStmt Then
        PreparedStatementsGet = ResultCode
    Else
        Set PreparedStatementsGet = PrepStmts
    End If
End Function


'''' ExecuteSetup prepares statement and/or binds parameters.
''''
'''' If non-blank SQLQuery is provided, a new query is prepared. If the
'''' statement has already been prerared, reset is called.
''''
'''' If query parameters are provided (as either an array or Name->Value
'''' dictionary map), bind is called.
'''' Reference: https://www.sqlite.org/lang_expr.html#varparam
''''
'''' Args:
''''   SQLQuery (string):
''''     SQL Query. Required parameter, but may be set to vbNullString.
''''
''''   QueryParameters (variant, optional, null):
''''     Parameters for a parameterized query.
''''
'''' Returns:
''''   SQLITE_OK, if successful; if not, other SQLiteResultCodes.
''''
'''' Raises:
''''   InvalidParameterErr
''''     If SQLQuery = vbNullString and either
''''       - the statement is not prepared,
''''       - QueryParameters is Null/Empty, or
''''       - QueryParameters is not an array or dictionary, or
''''       - existing query is not parameterized.
''''
'@Description "Prepares statement and/or binds parameters."
Friend Function ExecuteSetup(ByVal SQLQuery As String, _
                    Optional ByVal QueryParameters As Variant = Null, _
                    Optional ByVal UseDeclaredTypes As Boolean = False _
                    ) As SQLiteResultCodes
Attribute ExecuteSetup.VB_Description = "Prepares statement and/or binds parameters."
    Dim ResultCode As SQLiteResultCodes
    Dim ErrMsg As String
    '@Ignore AssignmentNotUsed: False positive
    ErrMsg = vbNullString
    Dim LenSQL As Long
    LenSQL = Len(SQLQuery)
    If LenSQL = 0 Then
        If IsNull(QueryParameters) Or IsEmpty(QueryParameters) Then
            ErrMsg = "Either query or parameters must be provided."
        ElseIf Not (VarType(QueryParameters) > vbArray Or _
                    TypeName(QueryParameters) = "Dictionary") Then
            ErrMsg = "QueryParameters must be an array or dictionary."
        ElseIf this.StmtHandle = 0 Then
            ErrMsg = "Blank query has been provided to an unprepared statement."
        End If
        Guard.ExpressionErr Len(ErrMsg) = 0, ErrNo.InvalidParameterErr, _
            MODULE_NAME & "/ExecuteSetup", ErrMsg
    End If
    this.PrepareOnly = False
    
    If LenSQL > 0 Then
        Dim SubQueries As Variant
        SubQueries = SQLlib.SplitMultiStatement(SQLQuery)
        ErrMsg = "Multistatement queries are only supported via ExecuteNonQuery."
        Guard.ExpressionErr UBound(SubQueries) = 0, ErrNo.InvalidParameterErr, _
            MODULE_NAME & "/ExecuteSetup", ErrMsg
        
        If this.StmtHandle <> 0 Then
            ResultCode = Finalize
            If ResultCode <> SQLITE_OK Then GoTo EXIT_FUNCTION:
        End If
        ResultCode = Prepare16V2(SQLQuery)
        If ResultCode <> SQLITE_OK Then GoTo EXIT_FUNCTION:
            
        '''' Retrieves ColumnCount and sets it on DbExec
        SQLiteCMeta.ColumnCountLet Me
        If UseDeclaredTypes Then
            '''' Retrieve parameter metadata
            ResultCode = SQLiteCMeta(Me).TableMetaCollect()
            If ResultCode <> SQLITE_OK Then
                Debug.Print "Failed to retrieve table metadata."
                GoTo EXIT_FUNCTION:
            End If
        Else
            this.DbExec.MetaLoaded = False
        End If
        
        '''' If query is parameterized, but no QueryParameters provided, do not
        '''' execute the statement (expect parameters in subsequent queries).
        this.PrepareOnly = this.ParameterCount > 0 And _
            (IsEmpty(QueryParameters) Or IsNull(QueryParameters))
        If this.PrepareOnly Then ResultCode = SQLITE_DONE
    Else
        '''' LenSQL = 0 And this.StmtHandle <> 0 - the only remaining option,
        '''' because <both = 0> option raises an error above.
        ResultCode = Reset
        If ResultCode <> SQLITE_OK Then
            ResultCode = Finalize
            If ResultCode = SQLITE_OK Then
                ResultCode = ExecuteSetup(SQLQueryOriginal, QueryParameters)
            End If
            GoTo EXIT_FUNCTION:
        End If
    End If
    
    If VarType(QueryParameters) > vbArray Or _
       TypeName(QueryParameters) = "Dictionary" Then
        '''' this.DbPara is set in Prepare16V2
        ResultCode = this.DbPara.BindDictOrArray(QueryParameters)
        If ResultCode = SQLITE_OK Then
            Logger.Logg "Parameters are bound.", , DEBUGLEVEL_INFO
        Else
            ResultCode = this.DbConn.ErrInfoRetrieve
        End If
    End If
EXIT_FUNCTION:
    ExecuteSetup = ResultCode
End Function


'''' If null is passed in place of QueryParameters, assume that the command
'''' is not parameterized and can be executed via the exec API. In this case,
'''' the command is relayed to DbConn.ExecuteNonQueryPlain, which can execute
'''' commands containing multiple semicolon-separated SQL non-query statements.
''''
'@Description "Executes SQL statement not returning data."
Public Function ExecuteNonQuery(ByVal SQLQuery As String, _
                       Optional ByVal QueryParameters As Variant = Empty, _
                       Optional ByRef AffectedRows As Long = -2) As SQLiteResultCodes
Attribute ExecuteNonQuery.VB_Description = "Executes SQL statement not returning data."
    Dim ResultCode As SQLiteResultCodes
    If IsNull(QueryParameters) Then
        ResultCode = this.DbConn.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
        Exit Function
    End If
    
    If AffectedRows <> -2 Then AffectedRows = this.DbConn.ChangesCount(True)
    
    Dim SubQueries As Variant
    SubQueries = SQLlib.SplitMultiStatement(SQLQuery)
    Dim SubQueryIndex As Long
    For SubQueryIndex = LBound(SubQueries) To UBound(SubQueries)
        ResultCode = ExecuteSetup(SubQueries(SubQueryIndex), QueryParameters, False)
        If ResultCode <> SQLITE_OK Then GoTo EXIT_FUNCTION:
        If Not this.PrepareOnly Then ResultCode = this.DbExec.ExecuteStepAPI
        If Not (ResultCode = SQLITE_OK Or ResultCode = SQLITE_DONE Or _
                ResultCode = SQLITE_ROW) Then GoTo EXIT_FUNCTION:
    Next SubQueryIndex
        
    If AffectedRows <> -2 Then
        AffectedRows = this.DbConn.ChangesCount(True) - AffectedRows
    End If
    
    this.AffectedRowsCount = AffectedRows
EXIT_FUNCTION:
    ExecuteNonQuery = ResultCode
End Function


'@Description "Executes SQL statement returning a scalar value (possibly Empty or an error)."
Public Function GetScalar(ByVal SQLQuery As String, _
                 Optional ByVal QueryParameters As Variant = Empty) As Variant
Attribute GetScalar.VB_Description = "Executes SQL statement returning a scalar value (possibly Empty or an error)."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters, False)
    If ResultCode = SQLITE_OK Then
        GetScalar = this.DbExec.GetScalar
    Else
        GetScalar = CVErr(ResultCode)
    End If
End Function


'@Description "Executes SQL statement returning a set of rows."
Public Function GetPagedRowSet(ByVal SQLQuery As String, _
                      Optional ByVal QueryParameters As Variant = Empty, _
                      Optional ByVal UseDeclaredTypes As Boolean = True) As Variant
Attribute GetPagedRowSet.VB_Description = "Executes SQL statement returning a set of rows."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters, UseDeclaredTypes)
    If ResultCode = SQLITE_OK Then
        GetPagedRowSet = this.DbExec.GetPagedRowSet()
        If this.DbConn.ErrInfo.ErrorCode = SQLITE_DONE Then Reset
    Else
        GetPagedRowSet = ResultCode
    End If
End Function


Public Function GetRowSet2D(ByVal SQLQuery As String, _
                   Optional ByVal QueryParameters As Variant = Empty, _
                   Optional ByVal UseDeclaredTypes As Boolean = True) As Variant
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters, UseDeclaredTypes)
    If ResultCode = SQLITE_OK Then
        GetRowSet2D = this.DbExec.GetRowSet2D()
        If this.DbConn.ErrInfo.ErrorCode = SQLITE_DONE Then Reset
    Else
        GetRowSet2D = ResultCode
    End If
End Function


Public Function GetRecordset(ByVal SQLQuery As String, _
                    Optional ByVal QueryParameters As Variant = Empty, _
                    Optional ByVal UseDeclaredTypes As Boolean = True) As SQLiteCRecordsetADO
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters, UseDeclaredTypes)
    If ResultCode <> SQLITE_OK Then
        Set GetRecordset = Nothing
        Exit Function
    End If
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = this.DbExec.GetFabRecordset()
    dbr.SQLQueryParamsSet QueryParameters
    If this.DbConn.ErrInfo.ErrorCode = SQLITE_DONE Then Reset
    Set GetRecordset = dbr
End Function


'''' Reference: https://www.sqlite.org/c3ref/expanded_sql.html
''''
'@Description "Returns SQL query supplied to the PSO factory with bound parameters expanded."
Public Property Get SQLQueryExpanded() As String
Attribute SQLQueryExpanded.VB_Description = "Returns SQL query supplied to the PSO factory with bound parameters expanded."
    If this.StmtHandle = 0 Then
        SQLQueryExpanded = vbNullString
        Exit Function
    End If
    #If VBA7 Then
        Dim SQLQueryPtr As LongPtr
    #Else
        Dim SQLQueryPtr As Long
    #End If
    SQLQueryPtr = sqlite3_expanded_sql(this.StmtHandle)
    Guard.ExpressionErr SQLQueryPtr <> 0, ErrNo.OutOfMemoryErr, _
        MODULE_NAME & "/SQLQueryExpanded", "SQLite DLL returned NULL pointer"
    Dim SQLQueryText As String
    SQLQueryText = UTFlib.StrFromUTF8Ptr(SQLQueryPtr)
    sqlite3_free SQLQueryPtr
    SQLQueryExpanded = SQLQueryText
End Property


'''' Reference: https://www.sqlite.org/c3ref/expanded_sql.html
''''
'@Description "Returns original (optionally normalized) SQL query supplied to the PSO factory."
Public Function SQLQueryOriginal(Optional ByVal Normalized As Boolean = False) As String
Attribute SQLQueryOriginal.VB_Description = "Returns original (optionally normalized) SQL query supplied to the PSO factory."
    If this.StmtHandle = 0 Then
        SQLQueryOriginal = vbNullString
        Exit Function
    End If
    #If VBA7 Then
        Dim SQLQueryPtr As LongPtr
    #Else
        Dim SQLQueryPtr As Long
    #End If
    If Normalized Then
        SQLQueryPtr = sqlite3_normalized_sql(this.StmtHandle)
    Else
        SQLQueryPtr = sqlite3_sql(this.StmtHandle)
    End If
    Guard.ExpressionErr SQLQueryPtr <> 0, ErrNo.OutOfMemoryErr, _
        MODULE_NAME & "/SQLQueryOriginal", "SQLite DLL returned NULL pointer"
    
    SQLQueryOriginal = UTFlib.StrFromUTF8Ptr(SQLQueryPtr)
End Function


'============================================================================='
'============================= ILiteADO INTERFACE ============================'
'============================================================================='


Private Property Get ILiteADO_MainDB() As String
    ILiteADO_MainDB = this.DbConn.DbPathName
End Property

'''' For SQLiteC, this is the same as MainDB
Private Property Get ILiteADO_ConnectionString() As String
    ILiteADO_ConnectionString = this.DbConn.DbPathName
End Property

Private Property Get ILiteADO_ClassName() As String
    ILiteADO_ClassName = MODULE_NAME
End Property

Private Function ILiteADO_OpenDb() As Long
    ILiteADO_OpenDb = this.DbConn.OpenDb
End Function

Private Function ILiteADO_CloseDb() As Long
    ILiteADO_CloseDb = this.DbConn.CloseDb
End Function

Private Property Get ILiteADO_ConnectionClosed() As Boolean
    ILiteADO_ConnectionClosed = (this.DbConn.DbHandle = 0)
End Property

Private Function ILiteADO_GetScalar(ByVal SQLQuery As String, _
                           Optional ByVal QueryParameters As Variant = Empty _
                           ) As Variant
    Dim dbc As SQLiteCConnection
    Set dbc = this.DbConn
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    If HandleConnection Then Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
        ErrNo.ConnectionOpenErr, MODULE_NAME & "/ILiteADO_GetScalar", _
        "Failed to open db."
    ILiteADO_GetScalar = GetScalar(SQLQuery, QueryParameters)
    If HandleConnection Then Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, _
        ErrNo.ConnectionCloseErr, MODULE_NAME & "/ILiteADO_GetScalar", _
        "Failed to close db."
End Function

Private Function ILiteADO_GetAdoRecordset(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As ADODB.Recordset
    Dim dbc As SQLiteCConnection
    Set dbc = this.DbConn
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    If HandleConnection Then Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
        ErrNo.ConnectionOpenErr, MODULE_NAME & "/ILiteADO_GetAdoRecordset", _
        "Failed to open db."
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = GetRecordset(SQLQuery, QueryParameters)
    Set ILiteADO_GetAdoRecordset = dbr.AdoRecordset
    If HandleConnection Then Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, _
        ErrNo.ConnectionCloseErr, MODULE_NAME & "/ILiteADO_GetAdoRecordset", _
        "Failed to close db."
End Function

Private Function ILiteADO_ExecuteNonQuery(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As Long
    Dim dbc As SQLiteCConnection
    Set dbc = this.DbConn
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    If HandleConnection Then Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
        ErrNo.ConnectionOpenErr, MODULE_NAME & "/ILiteADO_ExecuteNonQuery", _
        "Failed to open db."
    Dim AffectedRows As Long: AffectedRows = 0 '''' RD Workaround
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteNonQuery(SQLQuery, QueryParameters, AffectedRows)
    If ResultCode = SQLITE_OK Or ResultCode = SQLITE_DONE Then
        ILiteADO_ExecuteNonQuery = AffectedRows
    Else
        ILiteADO_ExecuteNonQuery = -1
    End If
    If HandleConnection Then Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, _
        ErrNo.ConnectionCloseErr, MODULE_NAME & "/ILiteADO_ExecuteNonQuery", _
        "Failed to close db."
End Function

Private Sub ILiteADO_Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = this.DbConn.Begin(Mode)
    Debug.Assert ResultCode = SQLITE_OK
End Sub

Private Sub ILiteADO_Commit()
    Dim ResultCode As SQLiteResultCodes
    ResultCode = this.DbConn.Commit
    Debug.Assert ResultCode = SQLITE_OK
End Sub

Private Sub ILiteADO_SavePoint(ByVal SavePointName As String)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = this.DbConn.SavePoint(SavePointName)
    Debug.Assert ResultCode = SQLITE_OK
End Sub

Private Sub ILiteADO_ReleasePoint(ByVal SavePointName As String)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = this.DbConn.ReleasePoint(SavePointName)
    Debug.Assert ResultCode = SQLITE_OK
End Sub

Private Sub ILiteADO_Rollback(Optional ByVal SavePointName As String = vbNullString)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = this.DbConn.Rollback(SavePointName)
    Debug.Assert ResultCode = SQLITE_OK
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
    Guard.EmptyString FeatureName
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
    '@Ignore AssignmentNotUsed:False positive
    Status = False
    With PragmaOptionsDict("compile_options")
        If .Exists(CompileOption) Then Status = CBool(.Item(CompileOption))
    End With
    ILiteADO_FeatureAvailable = Status
End Function
