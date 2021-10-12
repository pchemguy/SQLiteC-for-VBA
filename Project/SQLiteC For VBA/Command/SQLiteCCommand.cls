VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Executes statement and retrieves data"
'@Folder "SQLiteC For VBA.Command"
'@ModuleDescription "Executes statement and retrieves data"
'@PredeclaredId
'@Exposed

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

Const OutOfMemoryErr As Long = 7&

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
Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString

Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value
#Else
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As Long
' 0-based indices
Private Declare Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString

Private Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Private Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value
#End If

Private Type TSQLiteCCommand
    DbStmt As SQLiteCStatement
    ColumnCount As Long
    ColumnNames As Variant
    ColumnDbNames As Variant
    ColumnTableNames As Variant
    ColumnOriginNames As Variant
    ColumnTypes As Variant
    ColumnDeclaredTypes As Variant
    ColumnMap As Scripting.Dictionary
    FetchError As Boolean
    PageSize As Long
    RowCount As Long
End Type
Private this As TSQLiteCCommand
Public ResultRows As Variant


'@DefaultMember
Public Function Create(ByVal DbStmt As SQLiteCStatement) As SQLiteCCommand
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCCommand
    Set Instance = New SQLiteCCommand
    Instance.Init DbStmt
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbStmt As SQLiteCStatement)
    Guard.NullReference DbStmt
    Set this.DbStmt = DbStmt
    Set this.ColumnMap = New Scripting.Dictionary
    this.ColumnMap.CompareMode = TextCompare
    this.RowCount = 0
    this.PageSize = 10
End Sub


Public Property Get RollBackTx() As Boolean
    RollBackTx = this.FetchError
End Property


Public Function SQLiteTypeName(ByVal SQLiteType As SQLiteTypes) As String
    SQLiteTypeName = Array("INTEGER", "FLOAT", "TEXT", "BLOB", "NULL")(SQLiteType - 1)
End Function


Public Function GetColumnMeta() As Boolean
    Const OutOfMemoryErr As Long = 7
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    Dim ColumnCount As Long
    ColumnCount = sqlite3_column_count(hStmt)
    this.ColumnCount = ColumnCount
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        GetColumnMeta = False
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
    #If VBA7 Then
        Dim BufferPtr As LongPtr
    #Else
        Dim BufferPtr As Long
    #End If
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
    
    GetColumnMeta = True
    Exit Function
    
NULL_PTR:
    '''' Possible "OutOfMemoryErr" / "Failed to allocate memory for column name"
    GetColumnMeta = False
    Exit Function
End Function


Public Function GetScalar() As Variant
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_step(hStmt)
    Select Case ResultCode
        Case SQLITE_DONE
            this.FetchError = False
            GetScalar = Empty
            Exit Function
        Case Is <> SQLITE_ROW
            this.FetchError = True
            With this.DbStmt.DbConn.ErrorInfo
                .GetErr
                .PrintErr
            End With
            GetScalar = Empty
            Exit Function
    End Select
    
    '''' ResultCode = SQLITE_ROW
    Dim ColumnIndex As Long
    ColumnIndex = 0
    Dim ColumnType As SQLiteTypes
    ColumnType = sqlite3_column_type(hStmt, ColumnIndex)
    Dim Value As Variant
    Select Case ColumnType
        Case SQLITE_INTEGER
            Value = sqlite3_column_int64(hStmt, ColumnIndex)
            If VarType(Value) = vbCurrency Then Value = Value * 10000
        Case SQLITE_FLOAT
            Value = sqlite3_column_double(hStmt, ColumnIndex)
        Case SQLITE_TEXT
            Value = UTFlib.StrFromUTF8Ptr(sqlite3_column_text(hStmt, ColumnIndex))
        Case SQLITE_NULL
            Value = Null
        Case SQLITE_BLOB
            #If VBA7 Then
                Dim RemoteBufPtr As LongPtr
            #Else
                Dim RemoteBufPtr  As Long
            #End If
            RemoteBufPtr = sqlite3_column_blob(hStmt, ColumnIndex)
            Dim Length As Long
            Length = sqlite3_column_bytes(hStmt, ColumnIndex)
            Dim LocalBuffer() As Byte
            ReDim LocalBuffer(Length - 1)
            RtlCopyMemory VarPtr(LocalBuffer(0)), RemoteBufPtr, Length
            Value = LocalBuffer
    End Select
    GetScalar = Value
End Function
