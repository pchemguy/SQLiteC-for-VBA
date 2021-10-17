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
'@Folder "SQLiteC For VBA.Executor"
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

Public Enum SQLiteTypes
    SQLITE_INTEGER = 1&
    SQLITE_FLOAT = 2&
    SQLITE_TEXT = 3&
    SQLITE_BLOB = 4&
    SQLITE_NULL = 5&
End Enum

#If VBA7 Then
Private Declare PtrSafe Sub RtlCopyMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
#Else
Private Declare Sub RtlCopyMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
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

Private Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
#End If

Private Type TSQLiteCExecSQL
    DbStmt As SQLiteCStatement
    ColumnCount As Long
    ColumnNames As Variant
    ColumnDbNames As Variant
    ColumnTableNames As Variant
    ColumnOriginNames As Variant
    ColumnTypes As Variant
    ColumnDeclaredTypes As Variant
    ColumnMap As Scripting.Dictionary
    StepAPIError As Boolean
    PageCount As Long
    FilledPagesCount As Long
    PageSize As Long
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
    Set this.DbStmt = DbStmt
    Set this.ColumnMap = New Scripting.Dictionary
    this.ColumnMap.CompareMode = TextCompare
    this.RowCount = 0
    this.PageCount = 10
    this.PageSize = 100
End Sub


Public Function SQLiteTypeName(ByVal SQLiteType As SQLiteTypes) As String
    SQLiteTypeName = Array("INTEGER", "FLOAT", "TEXT", "BLOB", "NULL")(SQLiteType - 1)
End Function


Public Property Get FilledPagesCount() As Long
    FilledPagesCount = this.FilledPagesCount
End Property


Public Property Get RowCount() As Long
    RowCount = this.RowCount
End Property


'''' Reference: https://www.sqlite.org/c3ref/column_count.html
'''' Reference: https://www.sqlite.org/c3ref/column_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_database_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_decltype.html
'''' Reference: https://www.sqlite.org/c3ref/column_blob.html
''''
'@Description "Retrieves columns metadata."
Public Function GetColumnMeta() As SQLiteResultCodes
Attribute GetColumnMeta.VB_Description = "Retrieves columns metadata."
    #If VBA7 Then
        Dim hStmt As LongPtr
        Dim BufferPtr As LongPtr
    #Else
        Dim hStmt As Long
        Dim BufferPtr As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    If hStmt = 0 Then
        GetColumnMeta = SQLITE_ERROR
        Exit Function
    End If
    Dim ColumnCount As Long
    ColumnCount = GetColumnCount
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        GetColumnMeta = SQLITE_ERROR
        Exit Function
    End If
    Dim ColumnNames() As String
    Dim ColumnTypes() As SQLiteTypes
    Dim ColumnDbNames() As String
    Dim ColumnTableNames() As String
    Dim ColumnOriginNames() As String
    Dim ColumnDeclaredTypes() As String
    ReDim ColumnNames(0 To ColumnCount - 1)
    ReDim ColumnTypes(0 To ColumnCount - 1)
    ReDim ColumnDbNames(0 To ColumnCount - 1)
    ReDim ColumnTableNames(0 To ColumnCount - 1)
    ReDim ColumnOriginNames(0 To ColumnCount - 1)
    ReDim ColumnDeclaredTypes(0 To ColumnCount - 1)
    
    this.ColumnMap.RemoveAll
    Dim ColumnIndex As Long
    For ColumnIndex = 0 To ColumnCount - 1
        BufferPtr = sqlite3_column_name(hStmt, ColumnIndex)
        If BufferPtr = 0 Then GoTo NULL_PTR:
        ColumnNames(ColumnIndex) = UTFlib.StrFromUTF8Ptr(BufferPtr)
        
        BufferPtr = sqlite3_column_database_name(hStmt, ColumnIndex)
        If BufferPtr = 0 Then GoTo NULL_PTR:
        ColumnDbNames(ColumnIndex) = UTFlib.StrFromUTF8Ptr(BufferPtr)
        
        BufferPtr = sqlite3_column_table_name(hStmt, ColumnIndex)
        If BufferPtr = 0 Then GoTo NULL_PTR:
        ColumnTableNames(ColumnIndex) = UTFlib.StrFromUTF8Ptr(BufferPtr)
        
        BufferPtr = sqlite3_column_origin_name(hStmt, ColumnIndex)
        If BufferPtr = 0 Then GoTo NULL_PTR:
        ColumnOriginNames(ColumnIndex) = UTFlib.StrFromUTF8Ptr(BufferPtr)
        
        BufferPtr = sqlite3_column_decltype(hStmt, ColumnIndex)
        If BufferPtr = 0 Then GoTo NULL_PTR:
        ColumnDeclaredTypes(ColumnIndex) = UTFlib.StrFromUTF8Ptr(BufferPtr)
        
        ColumnTypes(ColumnIndex) = sqlite3_column_type(hStmt, ColumnIndex)
        this.ColumnMap(ColumnNames(ColumnIndex)) = ColumnIndex
    Next ColumnIndex
    With this
        .ColumnNames = ColumnNames
        .ColumnTypes = ColumnTypes
        .ColumnDbNames = ColumnDbNames
        .ColumnTableNames = ColumnTableNames
        .ColumnOriginNames = ColumnOriginNames
        .ColumnDeclaredTypes = ColumnDeclaredTypes
    End With
    
    GetColumnMeta = SQLITE_OK
    Exit Function
    
NULL_PTR:
    '''' Possible "OutOfMemoryErr" / "Failed to allocate memory for column name"
    GetColumnMeta = this.DbStmt.DbConnection.ErrInfoRetrieve
    Exit Function
End Function


'''' Reference: https://www.sqlite.org/c3ref/step.html
''''
Public Function Execute() As SQLiteResultCodes
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_step(this.DbStmt.StmtHandle)
    Select Case ResultCode
        Case SQLITE_DONE, SQLITE_OK, SQLITE_ROW
            this.StepAPIError = False
        Case Else
            this.StepAPIError = True
            ResultCode = this.DbStmt.DbConnection.ErrInfoRetrieve
    End Select
    Execute = ResultCode
End Function


'@Description "Retrieves a scalar value"
Public Function GetScalar() As Variant
Attribute GetScalar.VB_Description = "Retrieves a scalar value"
    If this.DbStmt.StmtHandle = 0 Then
        GetScalar = CVErr(SQLITE_ERROR)
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
    ResultCode = Execute
    If ResultCode = SQLITE_ROW Then
        GetScalar = GetColumnValue
    Else
        GetScalar = CVErr(ResultCode)
    End If
End Function


Public Function GetColumnCount() As Long
    Dim ColumnCount As Long
    ColumnCount = sqlite3_column_count(this.DbStmt.StmtHandle)
    this.ColumnCount = ColumnCount
    GetColumnCount = ColumnCount
End Function


'''' Reference: https://www.sqlite.org/c3ref/column_blob.html
''''
Public Function GetColumnValue(Optional ByVal ColumnIndex As Long = 0) As Variant
    #If VBA7 Then
        Dim hStmt As LongPtr
        Dim RemoteBufPtr As LongPtr
    #Else
        Dim hStmt As Long
        Dim RemoteBufPtr As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    
    Dim ColumnType As SQLiteTypes
    ColumnType = sqlite3_column_type(hStmt, ColumnIndex)
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
                RtlCopyMemory VarPtr(LocalBuffer(0)), RemoteBufPtr, Length
                Value = LocalBuffer
            Else
                Value = CVErr(this.DbStmt.DbConnection.ErrInfoRetrieve)
            End If
    End Select
    GetColumnValue = Value
End Function


Public Function GetRow(ByVal ColumnCount As Long) As Variant
    Dim ColumnIndex As Long
    Dim RowValues As Variant
    ReDim RowValues(0 To ColumnCount - 1)
    For ColumnIndex = 0 To ColumnCount - 1
        RowValues(ColumnIndex) = GetColumnValue(ColumnIndex)
        If IsError(RowValues(ColumnIndex)) Then
            RowValues = RowValues(ColumnIndex)
            Exit Function
        End If
    Next ColumnIndex
    GetRow = RowValues
End Function


Public Function GetRowSet() As Variant
    If this.DbStmt.StmtHandle = 0 Then
        GetRowSet = CVErr(SQLITE_ERROR)
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
        
    Dim ColumnCount As Long
    ColumnCount = GetColumnCount
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        GetRowSet = CVErr(SQLITE_ERROR)
        Exit Function
    End If
    Dim Page As Variant
    Dim RowValues As Variant
    Dim Pages As Variant: Pages = Empty '''' RDVBA workaround
    ReDim Pages(0 To this.PageCount - 1)
    Dim RowCount As Long
    RowCount = 0
    this.FilledPagesCount = 0
    Dim RowIndex As Long
    Dim ColumnIndex As Long
    Dim PageIndex As Long
    For PageIndex = 0 To this.PageCount - 1
        Page = Empty
        ReDim Page(0 To this.PageSize - 1, 0 To ColumnCount - 1)
        For RowIndex = 0 To this.PageSize - 1
            ResultCode = Execute
            Select Case ResultCode
                Case SQLITE_ROW
                    RowValues = GetRow(ColumnCount)
                    If IsError(RowValues) Then
                        GetRowSet = RowValues
                        Exit Function
                    End If
                    For ColumnIndex = 0 To ColumnCount - 1
                        Page(RowIndex, ColumnIndex) = RowValues(ColumnIndex)
                    Next ColumnIndex
                    RowCount = RowCount + 1
                Case SQLITE_DONE
                    Pages(PageIndex) = Page
                    this.RowCount = RowCount
                    ResultCode = this.DbStmt.DbConnection.ErrInfoRetrieve
                    GetRowSet = Pages
                    Exit Function
                Case Else
                    GetRowSet = CVErr(ResultCode)
                    Exit Function
            End Select
        Next RowIndex
        Pages(PageIndex) = Page
        this.FilledPagesCount = this.FilledPagesCount + 1
    Next PageIndex
    this.RowCount = RowCount
    GetRowSet = Pages
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
