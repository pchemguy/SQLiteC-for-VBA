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
'@Folder "SQLite.C.Statement.ExecSQL"
'@ModuleDescription "Executes statement and retrieves data"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
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
Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
#If Win64 Then
    Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
#Else
    Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Currency
#End If
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
#Else
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As Long
' 0-based indices
Private Declare Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
#End If

Private Type TSQLiteCExecSQL
    DbStmt As SQLiteCStatement
    TableMeta() As SQLiteCColumnMeta
    ColumnCount As Long
    MetaLoaded As Boolean
    StepAPIError As Boolean
    PageCount As Long
    PageSize As Long
    FilledPagesCount As Long
    RowCount As Long
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
        .RowCount = 0
        .PageCount = 10
        .PageSize = 100
        .MetaLoaded = False
    End With
End Sub


Public Sub CleanUp()
    Set this.DbStmt = Nothing
End Sub


Public Property Get MetaLoaded() As Boolean
    MetaLoaded = this.MetaLoaded
End Property


Public Property Let MetaLoaded(ByVal Value As Boolean)
    this.MetaLoaded = Value
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


'''' Reference: https://www.sqlite.org/c3ref/step.html
''''
'@Description "Executes statement or moves the cursor forward to the next available row."
Friend Function ExecuteStepAPI() As SQLiteResultCodes
Attribute ExecuteStepAPI.VB_Description = "Executes statement or moves the cursor forward to the next available row."
    With this.DbStmt
        If .StmtHandle = 0 Then Err.Raise ErrNo.StatementNotPreparedErr, _
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


'''' Reference: https://www.sqlite.org/c3ref/column_count.html
''''
'@Description "Retrieves number of columns returned by the prepared statement."
Public Function ColumnCountAPI() As Long
Attribute ColumnCountAPI.VB_Description = "Retrieves number of columns returned by the prepared statement."
    FixGuard.StmtNotPrepared this.DbStmt, "SQLiteCExecSQL/ColumnCountAPI"
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
'''' Reference: https://www.sqlite.org/c3ref/c_blob.html
''''
'@Description "Retrieves a single column value from currently available row."
Friend Function GetColumnValueAPI( _
            Optional ByVal ColumnIndex As Long = 0, _
            Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
Attribute GetColumnValueAPI.VB_Description = "Retrieves a single column value from currently available row."
    FixGuard.StmtNotPrepared this.DbStmt, "SQLiteCExecSQL/GetColumnValueAPI"
    #If VBA7 Then
        Dim hStmt As LongPtr
        Dim RemoteBufPtr As LongPtr
    #Else
        Dim hStmt As Long
        Dim RemoteBufPtr As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    
    '''' Must query value type API to check for NULL.
    '''' This call is not required for non-nullable columns if declared type is
    '''' used. Additional "if" is required, however, which should impose a slight
    '''' penalty in case when most columns are nullable and require this null-check.
    '''' The caller is responsible for querying/populating the metadata.
    '''' For calculated columns, AffinityType is not defined and should not be used.
    Dim ColumnType As SQLiteDataType
    ColumnType = sqlite3_column_type(hStmt, ColumnIndex)
    If this.MetaLoaded And ColumnType <> SQLITE_NULL And UseDeclaredTypes Then
        Dim AffinityType As SQLiteDataType
        AffinityType = this.TableMeta(ColumnIndex).AffinityType
        If AffinityType > SQLITE_NONE Then ColumnType = AffinityType
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


'@Description "Retrieves currently available row."
Public Function GetRow(ByVal ColumnCount As Long, _
              Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
Attribute GetRow.VB_Description = "Retrieves currently available row."
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


'''' Since SQLite provides a forward-only cursor and does not provide information
'''' about the total number of rows returned by the query, this is the primary
'''' row set query interface. At the start, a variant 1D 0-based Pages array is
'''' ReDim'ed, and its size is set to PageCount. As rows are retrieved from the
'''' database, they are saved as variant 1D 0-based array elements with in a
'''' variant 0-based 1D Page array. Once the Page array is filled to PageSize,
'''' it is saved into the next available slot in the Pages array. At the end,
'''' the Pages array is retuned to the caller, .FilledPagesCount is set to the
'''' number of fully filled pages, and .RowCount is set to the total number of
'''' rows retrieved (neither is stricty necessary, as IsEmpty(Pages(PageIndex))
'''' and IsEmpty(Page(RowIndex)) can be used to as While Loop conditions,
'''' replacing the currently used For Loops).
''''
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
        Dim dbsm As SQLiteCMeta
        Set dbsm = SQLiteCMeta(this.DbStmt)
        ResultCode = dbsm.TableMetaCollect()
        If ResultCode <> SQLITE_OK Then
            Debug.Print "Failed to retrieve table metadata."
            GetPagedRowSet = CVErr(ResultCode)
            Exit Function
        End If
        Set dbsm = Nothing
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


'''' This method is a wrapper around GetPagedRowSet (used to retrieve the data)
'''' reshaping the paged data set as a plain 0-based 2D array.
''''
'@Description "Shapes/flattens the data returning a 2D variant array."
Public Function GetRowSet2D(Optional ByVal UseDeclaredTypes As Boolean = False) As Variant
Attribute GetRowSet2D.VB_Description = "Shapes/flattens the data returning a 2D variant array."
    Dim PagedRowSet As Variant
    PagedRowSet = GetPagedRowSet(UseDeclaredTypes)
    If IsError(PagedRowSet) Then
        GetRowSet2D = PagedRowSet
        Exit Function
    ElseIf this.RowCount = 0 Then
        GetRowSet2D = Empty
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


'''' This method is a wrapper around GetPagedRowSet (used to retrieve the data)
'''' reshaping the paged data set. It fabricates an ADODB.Recordset object and
'''' populates it with retrieved data. This way, SQLiteC package can implement
'''' the ADODB based ILiteADO interface.
''''
'@Description "Shapes/flattens the data returning a fabricated ADODB.Recordset object."
Public Function GetFabRecordset(Optional ByVal UseDeclaredTypes As Boolean = False _
                               ) As SQLiteCRecordsetADO
Attribute GetFabRecordset.VB_Description = "Shapes/flattens the data returning a fabricated ADODB.Recordset object."
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
