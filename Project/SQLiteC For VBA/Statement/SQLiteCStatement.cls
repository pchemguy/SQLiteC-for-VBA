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
    '@Folder "SQLiteC For VBA.Statement"
'@ModuleDescription "Manages the lifecycle of the SQLite3 Prerpared Statement Object (PSO)."
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess

''''======================================================================''''
'''' Statement
'''' https://www.sqlite.org/c3ref/prepare.html
'''' https://www.sqlite.org/c3ref/reset.html
'''' https://www.sqlite.org/c3ref/finalize.html
''''======================================================================''''

Option Explicit

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
    AffectedRowsCount As Long
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
End Sub


Private Sub Class_Terminate()
    '@Ignore FunctionReturnValueDiscarded
    If this.StmtHandle <> 0 Then Finalize
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


'''' Reference: https://www.sqlite.org/c3ref/prepare.html
''''
'@Ignore UseMeaningfulName
'@Description "Generates SQLite PSO."
Friend Function Prepare16V2(ByVal SQLQuery As String) As SQLiteResultCodes
Attribute Prepare16V2.VB_Description = "Generates SQLite PSO."
    If this.DbConn.DbHandle = 0 Then Err.Raise ConnectionNotOpenedErr, _
        "SQLiteCStatement", "Database connection is not opened!"
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
                Err.Raise ErrNo.InvalidParameterErr, "SQLiteCStatement", _
                          "Failed to instantiate SQLiteCParameters."
            End If
        End If
    End If
    Prepare16V2 = ResultCode
End Function


'''' Reset should be called after execution before the statement object can be
'''' reused (for example, with a new set of bound parameters).
''''
'''' Reference: https://www.sqlite.org/c3ref/reset.html
''''
'@Description "Resets SQLite PSO after execution."
Friend Function Reset() As SQLiteResultCodes
Attribute Reset.VB_Description = "Resets SQLite PSO after execution."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_reset(this.StmtHandle)
    If ResultCode <> SQLITE_OK Then ResultCode = this.DbConn.ErrInfoRetrieve
    Reset = ResultCode
End Function


'''' sqlite3_finalize should only return an error if the the most recent
'''' evaluation of the statement failed. At the same time, if Prepare16V2
'''' returns an error, StmtHandle is set to 0 to prevent an error here. The
'''' only remaining documented possibility is an error generated in course of
'''' automatic recompilation (https://www.sqlite.org/c3ref/prepare.html).
''''
'''' Reference: https://www.sqlite.org/c3ref/finalize.html
''''
'@Description "Destructs SQLite PSO."
Friend Function Finalize() As SQLiteResultCodes
Attribute Finalize.VB_Description = "Destructs SQLite PSO."
    If Not this.DbPara Is Nothing Then this.DbPara.BindClear
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_finalize(this.StmtHandle)
    If ResultCode = SQLITE_OK Then
        this.StmtHandle = 0
        Debug.Print "Statement is finalized"
    Else
        ResultCode = this.DbConn.ErrInfoRetrieve
    End If
    Finalize = ResultCode
End Function


'''' Args:
''''   DbConn (SQLiteCConnection):
''''     Connection object
''''
''''   FinStmt (optional, boolean, true):
''''     If true, finalize statements
''''
'''' Reference: https://www.sqlite.org/c3ref/stmt_busy.html
'''' Reference: https://www.sqlite.org/c3ref/next_stmt.html
''''
'''' @ClassMethod "Call on the predeclared instance"
'@Description "Enumerates prepared statements with buys status."
Friend Function PreparedStatementsGet(ByVal DbConn As SQLiteCConnection, _
                             Optional ByVal FinStmt As Boolean = True) As Variant
Attribute PreparedStatementsGet.VB_Description = "Enumerates prepared statements with buys status."
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
                Debug.Print "Statement is finalized"
            Else
                ResultCode = DbConn.ErrInfoRetrieve
            End If
        End If
        PrepStmts(hStmt) = sqlite3_stmt_busy(hStmt)
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
                    Optional ByVal QueryParameters As Variant = Null _
                    ) As SQLiteResultCodes
Attribute ExecuteSetup.VB_Description = "Prepares statement and/or binds parameters."
    Dim ResultCode As SQLiteResultCodes
    Dim LenSQL As Long
    LenSQL = Len(SQLQuery)
    
    If LenSQL = 0 Then
        Dim ErrMsg As String
        ErrMsg = vbNullString
        If IsNull(QueryParameters) Or IsEmpty(QueryParameters) Then
            ErrMsg = "Either query or parameters must be provided."
        ElseIf Not (VarType(QueryParameters) > vbArray Or _
                    TypeName(QueryParameters)) = "Dictionary" Then
            ErrMsg = "QueryParameters must be an array or dictionary."
        ElseIf this.StmtHandle = 0 Then
            ErrMsg = "Blank query has been provided to an unprepared statement."
        Else
            If this.DbPara Is Nothing Then
                ErrMsg = "New query is not provided and existing query is " & _
                         "not parameterized."
            ElseIf this.DbPara.ParameterCount = 0 Then
                ErrMsg = "New query is not provided and existing query is " & _
                         "not parameterized."
            End If
        End If
        If Len(ErrMsg) > 0 Then Err.Raise ErrNo.InvalidParameterErr, _
                                          "SQLiteCStatement", ErrMsg
    End If
        
    If LenSQL > 0 Then
        If this.StmtHandle <> 0 Then
            ResultCode = Finalize
            If ResultCode <> SQLITE_OK Then
                ExecuteSetup = ResultCode
                Exit Function
            End If
        End If
        ResultCode = Prepare16V2(SQLQuery)
        If ResultCode <> SQLITE_OK Then
            ExecuteSetup = ResultCode
            Exit Function
        End If
    Else
        '''' LenSQL = 0 And this.StmtHandle <> 0 - the only remaining option,
        '''' because <both = 0> option raises an error above.
        ResultCode = Reset
        If ResultCode <> SQLITE_OK Then
            ResultCode = ExecuteSetup(SQLQueryOriginal, QueryParameters)
            ExecuteSetup = ResultCode
            Exit Function
        End If
    End If
    
    If VarType(QueryParameters) > vbArray Or _
       TypeName(QueryParameters) = "Dictionary" Then
        '''' this.DbPara is set in Prepare16V2
        ResultCode = this.DbPara.BindDictOrArray(QueryParameters)
        If ResultCode = SQLITE_OK Then
            Debug.Print "Parameters are bound."
        Else
            ResultCode = this.DbConn.ErrInfoRetrieve
        End If
    End If
    ExecuteSetup = ResultCode
End Function


Public Function ExecuteNonQuery(ByVal SQLQuery As String, _
            Optional ByRef AffectedRows As Long = -1, _
            Optional ByVal QueryParameters As Variant = Empty) As SQLiteResultCodes
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters)
    If ResultCode <> SQLITE_OK Then
        ExecuteNonQuery = ResultCode
        Exit Function
    End If
    
    If AffectedRows <> -1 Then AffectedRows = this.DbConn.ChangesCount(True)
    ResultCode = this.DbExec.Execute
    If ResultCode = SQLITE_OK Or ResultCode = SQLITE_DONE Or ResultCode = SQLITE_ROW Then
        If AffectedRows <> -1 Then
            AffectedRows = this.DbConn.ChangesCount(True) - AffectedRows
        End If
        ResultCode = Reset
        ExecuteNonQuery = ResultCode
    Else
        AffectedRows = -1
    End If
    If Not IsEmpty(QueryParameters) Then this.DbPara.BindClear
    this.AffectedRowsCount = AffectedRows
    ExecuteNonQuery = ResultCode
End Function


Public Function GetScalar(ByVal SQLQuery As String, _
                 Optional ByVal QueryParameters As Variant = Empty) As Variant
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters)
    If ResultCode = SQLITE_OK Then
        If Not IsEmpty(QueryParameters) Then this.DbPara.BindClear
        GetScalar = this.DbExec.GetScalar
    Else
        GetScalar = CVErr(ResultCode)
    End If
End Function


Public Function GetRowSet(ByVal SQLQuery As String, _
                 Optional ByVal QueryParameters As Variant = Empty) As Variant
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteSetup(SQLQuery, QueryParameters)
    If ResultCode = SQLITE_OK Then
        GetRowSet = this.DbExec.GetRowSet
        If this.DbConn.ErrorInfo.ErrorCode = SQLITE_DONE Then Reset
    Else
        GetRowSet = CVErr(ResultCode)
    End If
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
    If SQLQueryPtr = 0 Then
        Err.Raise OutOfMemoryErr, "SQLiteCStatement/SQLQueryExpanded", _
                  "SQLite DLL returned NULL pointer"
    End If
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
    If SQLQueryPtr = 0 Then
        Err.Raise OutOfMemoryErr, "SQLiteCStatement/SQLQuery", _
                  "SQLite DLL returned NULL pointer"
    End If
    SQLQueryOriginal = UTFlib.StrFromUTF8Ptr(SQLQueryPtr)
End Function
