VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCExecSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Executes statement and retrieves data"
'@Folder "SQLiteC For VBA.Statement"
'@ModuleDescription "Executes statement and retrieves data"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess

''''======================================================================''''
'''' Execute and retrieve
'''' https://www.sqlite.org/c3ref/exec.html
'''' https://www.sqlite.org/c3ref/step.html
'''' https://www.sqlite.org/c3ref/column_blob.html
'''' https://www.sqlite.org/c3ref/column_count.html
'''' https://www.sqlite.org/c3ref/column_name.html
'''' https://www.sqlite.org/c3ref/c_blob.html
''''======================================================================''''

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
#Else
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
#End If

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
' 0-based indices
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_table_column_metadata Lib "SQLite3" ( _
    ByVal hDb As LongPtr, ByVal zDbName As LongPtr, ByVal zTableName As LongPtr, ByVal zColumnName As LongPtr, _
    ByRef pzDataType As LongPtr, ByRef pzCollSeq As LongPtr, ByRef pNotNull As Long, _
    ByRef pPrimaryKey As Long, ByRef pAutoinc As Long) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
#Else
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As Long
' 0-based indices
Private Declare Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_table_column_metadata Lib "SQLite3" ( _
    ByVal hDb As Long, ByVal zDbName As Long, ByVal zTableName As Long, ByVal zColumnName As Long, _
    ByRef pzDataType As Long, ByRef pzCollSeq As Long, ByRef pNotNull As Long, _
    ByRef pPrimaryKey As Long, ByRef pAutoinc As Long) As SQLiteResultCodes
Private Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
#End If

Private Type TSQLiteCExecSQL
    DbStmt As SQLiteCStatement
    TableMeta() As SQLiteCColumnMeta
    ColumnMap As Scripting.Dictionary
    ColumnCount As Long
    MetaLoaded As Boolean
    AffinityMap() As Variant ' SQLiteType
    StepAPIError As Boolean
    PageCount As Long
    PageSize As Long
    FilledPagesCount As Long
    RowCount As Long
    '''' Meta API's fenced with the .UseTableMetadataAPI flag will fail if executed
    '''' against the PRAGMA-Table-functions SELECTs. Set it to false in such a case.
    UseTableMetaAPI As Boolean
End Type
Private this As TSQLiteCExecSQL


'@DefaultMember
Public Function Create(ByVal DbStmt As SQLiteCStatement) As SQLiteCExecSQL
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCExecSQL
    Set Instance = New SQLiteCExecSQL
    Instance.Init DbStmt
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbStmt As SQLiteCStatement)
    Guard.NullReference DbStmt
    With this
        Set .DbStmt = DbStmt
        Set .ColumnMap = New Scripting.Dictionary
        .ColumnMap.CompareMode = TextCompare
        .RowCount = 0
        .PageCount = 10
        .PageSize = 100
        .MetaLoaded = False
        .AffinityMap = Array(SQLITE_BLOB, SQLITE_TEXT, SQLITE_TEXT, _
                             SQLITE_INTEGER, SQLITE_FLOAT)
        .UseTableMetaAPI = True
    End With
End Sub


Public Function SQLiteTypeName(ByVal DataTypeCode As SQLiteType) As String
    SQLiteTypeName = Array("NONE", "INTEGER", "FLOAT", _
                           "TEXT", "BLOB", "NULL")(DataTypeCode)
End Function


Public Function SQLiteTypeAffinityName(ByVal DataTypeAffinityCode As SQLiteTypeAffinity) As String
    SQLiteTypeAffinityName = Array("NONE", "BLOB", "TEXT", "NUMERIC", _
                                   "INTEGER", "REAL")(DataTypeAffinityCode - &H40)
End Function


Public Property Get MetaLoaded() As Boolean
    MetaLoaded = this.MetaLoaded
End Property


Public Property Get TableMeta() As SQLiteCColumnMeta()
    TableMeta = this.TableMeta
End Property


Public Property Let TableMeta(ByRef Value() As SQLiteCColumnMeta)
    this.TableMeta = Value
End Property


'@Ignore ProcedureNotUsed
Public Property Get FilledPagesCount() As Long
    FilledPagesCount = this.FilledPagesCount
End Property


Public Property Get RowCount() As Long
    RowCount = this.RowCount
End Property


'@Ignore ProcedureNotUsed
Public Property Get UseTableMetaAPI() As Boolean
    UseTableMetaAPI = this.UseTableMetaAPI
End Property


Public Property Let UseTableMetaAPI(ByVal Value As Boolean)
    this.UseTableMetaAPI = Value
End Property


'@Ignore ProcedureNotUsed
Public Property Get PageCount() As Long
    PageCount = this.PageCount
End Property


Public Property Let PageCount(ByVal Value As Long)
    this.PageCount = Value
End Property


'@Ignore ProcedureNotUsed
Public Property Get PageSize() As Long
    PageSize = this.PageSize
End Property


Public Property Let PageSize(ByVal Value As Long)
    this.PageSize = Value
End Property


Friend Property Get AffinityMap() As Variant
    AffinityMap = this.AffinityMap
End Property


'''' Reference: https://www.sqlite.org/datatype3.html
''''            See section 3.1
''''
'@Description "Determines column type affinity."
Friend Function TypeAffinityFromDeclaredType(ByVal DeclaredType As String) As SQLiteTypeAffinity
Attribute TypeAffinityFromDeclaredType.VB_Description = "Determines column type affinity."
    Dim DeclType As String
    DeclType = UCase$(DeclaredType)
    If InStr(DeclType, "INT") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_INTEGER
    ElseIf InStr(DeclType, "CHAR") > 0 Or _
           InStr(DeclType, "CLOB") > 0 Or _
           InStr(DeclType, "TEXT") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_TEXT
    ElseIf InStr(DeclType, "BLOB") > 0 Or Len(DeclType) = 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_BLOB
    ElseIf InStr(DeclType, "REAL") > 0 Or _
           InStr(DeclType, "FLOA") > 0 Or _
           InStr(DeclType, "DOUB") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_REAL
    Else
        TypeAffinityFromDeclaredType = SQLITE_AFF_NUMERIC
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/table_column_metadata.html
'''' Reference: https://www.sqlite.org/c3ref/column_count.html
'''' Reference: https://www.sqlite.org/c3ref/column_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_database_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_decltype.html
'''' Reference: https://www.sqlite.org/c3ref/column_blob.html
''''
'''' For calculated columns, only the column_name API should work. Other APIs
'''' should return null pointer, without indicating an error otherwise.
''''
'@Description "Retrieves columns metadata."
Friend Function ColumnMetaAPI(ByRef ColumnInfo As SQLiteCColumnMeta _
                             ) As SQLiteResultCodes
Attribute ColumnMetaAPI.VB_Description = "Retrieves columns metadata."
    Dim dbs As SQLiteCStatement
    Set dbs = this.DbStmt
    Dim dbc As SQLiteCConnection
    Set dbc = dbs.DbConnection
    
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = dbs.StmtHandle
    If hStmt = 0 Then Err.Raise StatementNotPreparedErr, "SQLiteCExecSQL", _
        "Statement is not prepared."
    
    Dim ResultCode As SQLiteResultCodes
    
    With ColumnInfo
        If .Initialized <> -1 Then Err.Raise ErrNo.InvalidParameterErr, _
            "SQLiteCExecSQL", "Set .ColumnIndex and .Initialized=-1 before the call!"

        Dim ColumnIndex As Long
        ColumnIndex = .ColumnIndex
        
        Dim ColumnNamePtr As Variant
        ColumnNamePtr = sqlite3_column_name(hStmt, ColumnIndex)
        If ColumnNamePtr = 0 Then GoTo META_ERROR:
        Dim AliasName As String
        AliasName = UTFlib.StrFromUTF8Ptr(ColumnNamePtr)
        .Name = AliasName
        this.ColumnMap(.Name) = ColumnIndex
        AliasName = Replace(LCase$(AliasName), "_", vbNullString)
        .RowId = (AliasName = "rowid" Or AliasName = "oid")
        
        '''' column_data_type API returns value type. It should return SQLITE_NULL,
        '''' if SQLITE_ROW has not been returned by the step API.
        .DataType = sqlite3_column_type(hStmt, ColumnIndex)
        
        Dim DatabaseNamePtr As Variant
        DatabaseNamePtr = sqlite3_column_database_name(hStmt, ColumnIndex)
        If DatabaseNamePtr <> 0 Then .DbName = UTFlib.StrFromUTF8Ptr(DatabaseNamePtr)
        
        Dim TableNamePtr As Variant
        TableNamePtr = sqlite3_column_table_name(hStmt, ColumnIndex)
        If TableNamePtr <> 0 Then .TableName = UTFlib.StrFromUTF8Ptr(TableNamePtr)
        
        ColumnNamePtr = sqlite3_column_origin_name(hStmt, ColumnIndex)
        If ColumnNamePtr <> 0 Then .OriginName = UTFlib.StrFromUTF8Ptr(ColumnNamePtr)
        
        '''' Meta API's fenced with the .UseTableMetadataAPI flag will fail if executed
        '''' against the PRAGMA-Table-functions SELECTs. Set it to false in such a case.
        If this.UseTableMetaAPI Then
            Dim DataTypePtr As Variant
            DataTypePtr = sqlite3_column_decltype(hStmt, ColumnIndex)
            If DataTypePtr <> 0 Then
                .DeclaredTypeC = UTFlib.StrFromUTF8Ptr(DataTypePtr)
                .Affinity = TypeAffinityFromDeclaredType(.DeclaredTypeC)
                .AffinityType = this.AffinityMap(.Affinity - SQLITE_AFF_NONE - 1)
            Else
                '''' SQLITE_RANGE is expected for a calculated column, such as "count(*)"
                '''' Only return an error, if this is not the case; otherwise, continue.
                dbc.ErrorInfo.GetErr
                If dbc.ErrorInfo.ErrorCode <> SQLITE_RANGE Then GoTo META_ERROR:
            End If

            '''' For queries against PRAGMA functions, table_column_metadata
            '''' should return SQLITE_ERROR (no such table column...).
            ''''
            '''' Make sure that relevant previous APIs did not return null pointers.
            '''' This issue occurs for calculated columns, but basic armor would cut
            '''' off this information for all columns.
            ''''
            If DatabaseNamePtr <> 0 And TableNamePtr <> 0 And ColumnNamePtr <> 0 Then
                Dim NotNull As Long
                NotNull = False
                Dim PrimaryKey As Long
                PrimaryKey = False
                Dim AutoIncrement As Long
                AutoIncrement = False
                Dim CollationPtr As Variant: CollationPtr = 0 '''' RD workaround
                ResultCode = sqlite3_table_column_metadata(dbc.DbHandle, _
                    DatabaseNamePtr, TableNamePtr, ColumnNamePtr, _
                    DataTypePtr, CollationPtr, NotNull, PrimaryKey, AutoIncrement)
                If ResultCode = SQLITE_OK And DataTypePtr <> 0 And CollationPtr <> 0 Then
                    .DeclaredTypeT = UTFlib.StrFromUTF8Ptr(DataTypePtr)
                    .Collation = UTFlib.StrFromUTF8Ptr(CollationPtr)
                    .NotNull = CBool(NotNull)
                    .PrimaryKey = CBool(PrimaryKey)
                    .AutoIncrement = CBool(AutoIncrement)
                Else
                    ResultCode = dbc.ErrInfoRetrieve
                End If
            End If
        End If
    End With
    
    ColumnMetaAPI = ResultCode
    Exit Function
    
META_ERROR:
    dbc.ErrorInfo.GetErr
    ResultCode = dbc.ErrorInfo.ErrorCode
    If ResultCode = SQLITE_OK Then
        ResultCode = SQLITE_ERROR
        If DatabaseNamePtr = 0 Then Debug.Print "ERROR:" & vbLf & "SQLite " & _
            "MetaAPI returned a null pointer without an error." & vbLf & _
            "Do you have a calculated column, such as 'count(*)'?"
    Else
        ResultCode = dbc.ErrInfoRetrieve
    End If
    ColumnMetaAPI = ResultCode
    Exit Function
End Function


'@Description "Collects table metadata."
Public Function TableMetaCollect() As SQLiteResultCodes
Attribute TableMetaCollect.VB_Description = "Collects table metadata."
    Dim ColumnCount As Long
    ColumnCount = ColumnCountAPI
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        TableMetaCollect = SQLITE_ERROR
        Exit Function
    End If
    ReDim this.TableMeta(0 To ColumnCount - 1)
    Dim ResultCode As SQLiteResultCodes
    Dim ColumnIndex As Long
    For ColumnIndex = 0 To ColumnCount - 1
        this.TableMeta(ColumnIndex).ColumnIndex = ColumnIndex
        this.TableMeta(ColumnIndex).Initialized = -1
        '''' column_data_type API returns value type. If SQLITE_ROW has not been
        '''' returned by step API, column_data_type should return SQLITE_NULL
        ResultCode = ColumnMetaAPI(this.TableMeta(ColumnIndex))
        If ResultCode <> SQLITE_OK Then
            TableMetaCollect = ResultCode
            Exit Function
        End If
    Next ColumnIndex
    
    this.MetaLoaded = True
    TableMetaCollect = ResultCode
End Function


'''' Reference: https://www.sqlite.org/c3ref/step.html
''''
Friend Function ExecuteStepAPI() As SQLiteResultCodes
    With this.DbStmt
        If .StmtHandle = 0 Then Err.Raise StatementNotPreparedErr, _
                "SQLiteCExecSQL", "Statement is not prepared."
        Dim ResultCode As SQLiteResultCodes
        ResultCode = sqlite3_step(.StmtHandle)
        If ResultCode <> SQLITE_ROW Then
            If ResultCode <> SQLITE_DONE And ResultCode <> SQLITE_OK Then
                this.StepAPIError = True
                ResultCode = .DbConnection.ErrInfoRetrieve
            End If
        End If
    End With
    ExecuteStepAPI = ResultCode
End Function


Public Function ColumnCountAPI() As Long
    If this.DbStmt.StmtHandle = 0 Then Err.Raise StatementNotPreparedErr, _
        "SQLiteCExecSQL", "Statement is not prepared."
    this.ColumnCount = sqlite3_column_count(this.DbStmt.StmtHandle)
    ColumnCountAPI = this.ColumnCount
End Function


'@Description "Retrieves a scalar value"
Public Function GetScalar() As Variant
Attribute GetScalar.VB_Description = "Retrieves a scalar value"
    Dim ResultCode As SQLiteResultCodes
    ResultCode = ExecuteStepAPI
    If ResultCode = SQLITE_ROW Then
        GetScalar = GetColumnValueAPI
    Else
        GetScalar = CVErr(ResultCode)
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/column_blob.html
''''
Friend Function GetColumnValueAPI( _
            Optional ByVal ColumnIndex As Long = 0, _
            Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
    #If VBA7 Then
        Dim hStmt As LongPtr
        Dim RemoteBufPtr As LongPtr
    #Else
        Dim hStmt As Long
        Dim RemoteBufPtr As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    If hStmt = 0 Then Err.Raise StatementNotPreparedErr, "SQLiteCExecSQL", _
        "Statement is not prepared."
    
    Dim ColumnType As SQLiteType
    '''' Must query value type API to check for NULL.
    '''' This call is not required for non-nullable columns if declared type is
    '''' used. Additional "if" is required, however, which should impose a slight
    '''' penalty in case when most columns are nullable and require this null-check.
    '''' The caller is responsible for querying/populating the metadata.
    ColumnType = sqlite3_column_type(hStmt, ColumnIndex)
    If this.MetaLoaded Then
        Dim Affinity As SQLiteTypeAffinity
        Affinity = this.TableMeta(ColumnIndex).Affinity - SQLITE_AFF_NONE - 1
        If ColumnType <> SQLITE_NULL And UseDeclaredTypes And Affinity >= 0 Then
            ColumnType = this.AffinityMap(Affinity)
        End If
    End If
    
    Dim Value As Variant
    Select Case ColumnType
        Case SQLITE_INTEGER
            Value = sqlite3_column_int64(hStmt, ColumnIndex)
            If VarType(Value) = vbCurrency Then Value = Value * 10000
            If Abs(Value) < &H7FFFFFFF Then Value = CLng(Value)
        Case SQLITE_FLOAT
            Value = sqlite3_column_double(hStmt, ColumnIndex)
        Case SQLITE_TEXT
            RemoteBufPtr = sqlite3_column_text(hStmt, ColumnIndex)
            If RemoteBufPtr > 0 Then
                Value = UTFlib.StrFromUTF8Ptr(RemoteBufPtr)
            Else
                Value = CVErr(this.DbStmt.DbConnection.ErrInfoRetrieve)
            End If
        Case SQLITE_NULL
            Value = Null
        Case SQLITE_BLOB
            RemoteBufPtr = sqlite3_column_blob(hStmt, ColumnIndex)
            Dim Length As Long
            Length = sqlite3_column_bytes(hStmt, ColumnIndex)
            If Length > 0 Then
                Dim LocalBuffer() As Byte
                ReDim LocalBuffer(Length - 1)
                RtlMoveMemory VarPtr(LocalBuffer(0)), RemoteBufPtr, Length
                Value = LocalBuffer
            Else
                Value = CVErr(this.DbStmt.DbConnection.ErrInfoRetrieve)
            End If
    End Select
    GetColumnValueAPI = Value
End Function


Public Function GetRow(ByVal ColumnCount As Long, _
              Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
    Dim ColumnIndex As Long
    Dim RowValues As Variant
    ReDim RowValues(0 To ColumnCount - 1)
    For ColumnIndex = 0 To ColumnCount - 1
        RowValues(ColumnIndex) = GetColumnValueAPI(ColumnIndex, UseDeclaredTypes)
        If IsError(RowValues(ColumnIndex)) Then
            RowValues = RowValues(ColumnIndex)
            Exit Function
        End If
    Next ColumnIndex
    GetRow = RowValues
End Function


Public Function GetPagedRowSet(Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
    Dim ResultCode As SQLiteResultCodes
        
    Dim ColumnCount As Long
    ColumnCount = ColumnCountAPI
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        GetPagedRowSet = CVErr(SQLITE_ERROR)
        Exit Function
    End If
    
    If UseDeclaredTypes Then
    ResultCode = TableMetaCollect
        If ResultCode <> SQLITE_OK Then
            Debug.Print "Failed to retrieve table metadata."
            GetPagedRowSet = CVErr(ResultCode)
            Exit Function
        End If
    End If
    
    Dim Page As Variant
    Dim RowValues As Variant
    Dim Pages As Variant: Pages = Empty '''' RDVBA workaround
    ReDim Pages(0 To this.PageCount - 1)
    Dim RowCount As Long
    RowCount = 0
    this.FilledPagesCount = 0
    Dim RowIndex As Long
    Dim PageIndex As Long
    For PageIndex = 0 To this.PageCount - 1
        Page = Empty
        ReDim Page(0 To this.PageSize - 1)
        For RowIndex = 0 To this.PageSize - 1
            ResultCode = ExecuteStepAPI
            Select Case ResultCode
                Case SQLITE_ROW
                    RowValues = GetRow(ColumnCount, UseDeclaredTypes)
                    If IsError(RowValues) Then
                        GetPagedRowSet = RowValues
                        Exit Function
                    End If
                    Page(RowIndex) = RowValues
                    RowCount = RowCount + 1
                Case SQLITE_DONE
                    Pages(PageIndex) = Page
                    this.RowCount = RowCount
                    GetPagedRowSet = Pages
                    Exit Function
                Case Else
                    GetPagedRowSet = CVErr(ResultCode)
                    Exit Function
            End Select
        Next RowIndex
        Pages(PageIndex) = Page
        this.FilledPagesCount = this.FilledPagesCount + 1
    Next PageIndex
    this.RowCount = RowCount
    GetPagedRowSet = Pages
End Function


Public Function GetRowSet2D(Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
    Dim PagedRowSet As Variant
    PagedRowSet = GetPagedRowSet(UseDeclaredTypes)
    If IsError(PagedRowSet) Then
        GetRowSet2D = PagedRowSet
        Exit Function
    End If
    
    Dim RowSetArray() As Variant
    ReDim RowSetArray(0 To this.RowCount - 1, 0 To this.ColumnCount - 1)
    
    Dim RecordIndex As Long
    RecordIndex = 0
    Dim ColumnIndex As Long
    Dim RowIndex As Long
    Dim PageIndex As Long
    For PageIndex = 0 To this.FilledPagesCount - 1
        For RowIndex = 0 To this.PageSize - 1
            For ColumnIndex = 0 To this.ColumnCount - 1
                RowSetArray(RecordIndex, ColumnIndex) = PagedRowSet(PageIndex)(RowIndex)(ColumnIndex)
            Next ColumnIndex
            RecordIndex = RecordIndex + 1
        Next RowIndex
    Next PageIndex
    For RowIndex = 0 To this.RowCount - this.FilledPagesCount * this.PageSize - 1
        For ColumnIndex = 0 To this.ColumnCount - 1
            RowSetArray(RecordIndex, ColumnIndex) = PagedRowSet(PageIndex)(RowIndex)(ColumnIndex)
        Next ColumnIndex
        RecordIndex = RecordIndex + 1
    Next RowIndex
    GetRowSet2D = RowSetArray
End Function


Public Function GetFabRecordset(Optional ByVal UseDeclaredTypes As Boolean = False _
                               ) As SQLiteCRecordsetADO
    Dim PagedRowSet As Variant
    PagedRowSet = GetPagedRowSet(UseDeclaredTypes)
    If IsError(PagedRowSet) Then
        Set GetFabRecordset = Nothing
        Exit Function
    End If
    
    Dim FabRecordset As SQLiteCRecordsetADO
    Set FabRecordset = SQLiteCRecordsetADO(this.DbStmt)

    Dim ColumnIndex As Long
    Dim RowIndex As Long
    Dim PageIndex As Long
    With FabRecordset.AdoRecordset
        .PageSize = this.PageSize
        .CacheSize = this.PageSize
        .Open
        For PageIndex = 0 To this.FilledPagesCount - 1
            For RowIndex = 0 To this.PageSize - 1
                .AddNew
                For ColumnIndex = 0 To this.ColumnCount - 1
                    '@Ignore ImplicitDefaultMemberAccess
                    .Fields(ColumnIndex) = PagedRowSet(PageIndex)(RowIndex)(ColumnIndex)
                Next ColumnIndex
                .Update
            Next RowIndex
        Next PageIndex
        For RowIndex = 0 To this.RowCount - this.FilledPagesCount * this.PageSize - 1
            .AddNew
            For ColumnIndex = 0 To this.ColumnCount - 1
                '@Ignore ImplicitDefaultMemberAccess
                .Fields(ColumnIndex) = PagedRowSet(PageIndex)(RowIndex)(ColumnIndex)
            Next ColumnIndex
            .Update
        Next RowIndex
        .UpdateBatch
    End With
    
    Set GetFabRecordset = FabRecordset
End Function



'#If VBA7 Then
'Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
'Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
'Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
'Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
'Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value
'#Else
'Private Declare Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
'Private Declare Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
'Private Declare Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
'Private Declare Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
'Private Declare Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value
'#End If
