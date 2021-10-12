VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Manages SQLite connectivity"
'@Folder "SQLiteC For VBA.Connection"
'@ModuleDescription "Manages SQLite connectivity"
'@PredeclaredId
'@Exposed

''''======================================================================''''
'''' Connection
'''' https://www.sqlite.org/c3ref/open.html
'''' https://www.sqlite.org/c3ref/close.html
'''' https://www.sqlite.org/vfs.html
'''' https://www.sqlite.org/c3ref/changes.html
'''' https://www.sqlite.org/c3ref/total_changes.html
''''======================================================================''''

Option Explicit

Private Const VERBOUS As Boolean = True

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#If WIN64 Then
Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As LongLong
#Else
Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
#End If
'''' Open/close connection
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As SQLiteErrors
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, _
    ByVal iFlags As Long, ByVal zVfs As LongPtr) As SQLiteErrors ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
'''' Transactions
Private Declare PtrSafe Function sqlite3_txn_state Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSchema As LongPtr) As SQLiteTxnState
'''' Query wrapper
Private Declare PtrSafe Function sqlite3_exec Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSql As LongPtr, _
    ByVal Callback As LongPtr, ByVal pArg As LongPtr, ByRef pzErrMsg As LongPtr) As SQLiteResultCodes
'''' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
#Else
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
Private Declare Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
'''' Open/close connection
Private Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As SQLiteResultCodes ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, _
    ByVal zVfs As Long) As SQLiteResultCodes ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3" (ByVal hDb As Long) As Long
'''' Transactions
Private Declare Function sqlite3_txn_state Lib "SQLite3" (ByVal hDb As Long, ByVal zSchema As Long) As SQLiteTxnState
'''' Query wrapper
Private Declare Function sqlite3_exec Lib "SQLite3" (ByVal hDb As Long, ByVal zSql As Long, _
    ByVal Callback As Long, ByVal pArg As Long, ByRef pzErrMsg As Long) As SQLiteResultCodes
'''' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As Long) As Long
#End If

Public Enum SQLiteOpenFlags
    SQLITE_OPEN_READONLY = &H1&
    SQLITE_OPEN_READWRITE = &H2&
    SQLITE_OPEN_CREATE = &H4&
    SQLITE_OPEN_DELETEONCLOSE = &H8&
    SQLITE_OPEN_EXCLUSIVE = &H10&
    SQLITE_OPEN_AUTOPROXY = &H20&
    SQLITE_OPEN_URI = &H40&
    SQLITE_OPEN_MEMORY = &H80&
    SQLITE_OPEN_MAIN_DB = &H100&
    SQLITE_OPEN_TEMP_DB = &H200&
    SQLITE_OPEN_TRANSIENT_DB = &H400&
    SQLITE_OPEN_MAIN_JOURNAL = &H800&
    SQLITE_OPEN_TEMP_JOURNAL = &H1000&
    SQLITE_OPEN_SUBJOURNAL = &H2000&
    SQLITE_OPEN_SUPER_JOURNAL = &H4000&
    SQLITE_OPEN_NOMUTEX = &H8000&
    SQLITE_OPEN_FULLMUTEX = &H10000
    SQLITE_OPEN_SHAREDCACHE = &H20000
    SQLITE_OPEN_PRIVATECACHE = &H40000
    SQLITE_OPEN_WAL = &H80000
    SQLITE_OPEN_NOFOLLOW = &H1000000
    SQLITE_OPEN_DEFAULT = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
End Enum

Public Enum SQLiteTxnMode
    SQLITE_TXN_DEFERRED = 0&
    SQLITE_TXN_IMMEDIATE = 1&
    SQLITE_TXN_EXCLUSIVE = 2&
End Enum

Public Enum SQLiteTxnState
    SQLITE_TXN_NONE = 0&
    SQLITE_TXN_READ = 1&
    SQLITE_TXN_WRITE = 2&
End Enum

Private Type TSQLiteCConnection
    DbPathName As String
    #If VBA7 Then
        DbHandle As LongPtr
    #Else
        DbHandle As Long
    #End If
    ErrorInfo As SQLiteCErr
End Type
Private this As TSQLiteCConnection


'@DefaultMember
Public Function Create(ByVal DbPathName As String) As SQLiteCConnection
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCConnection
    Set Instance = New SQLiteCConnection
    Instance.Init DbPathName
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbPathName As String)
    Dim FilePathName As String
    FilePathName = DbPathName
    '''' TODO: DbPath validation
    ''''       This task should be performed by SQLiteDBVBA.LiteFSCheck class.
    ''''       This functionality will be integrated after the code is merged
    ''''       into the SQLiteDB VBA Library.
    this.DbPathName = FilePathName
    Set this.ErrorInfo = SQLiteCErr(Me)
End Sub


#If VBA7 Then
Public Property Get DbHandle() As LongPtr
#Else
Public Property Get DbHandle() As Long
#End If
    DbHandle = this.DbHandle
End Property


Public Property Get DbPathName() As String
    DbPathName = this.DbPathName
End Property


Public Property Get ErrorInfo() As SQLiteCErr
    Set ErrorInfo = this.ErrorInfo
End Property


Public Sub ErrInfoRetrieve()
    With this.ErrorInfo
        .GetErr
        .PrintErr
    End With
End Sub


Public Function Version(Optional ByVal Numeric As Boolean = True) As Variant
    If Numeric Then
        Version = sqlite3_libversion_number()
    Else
        Version = UTFlib.StrFromUTF8Ptr(sqlite3_libversion())
    End If
End Function


Public Function VersionI64(Optional ByVal Numeric As Boolean = True) As Variant
    Dim VerNum As Variant
    VerNum = sqlite3_libversion_number_i64()
    If VarType(VerNum) = vbCurrency Then VerNum = VerNum * 10000
    VersionI64 = VerNum
End Function


Public Sub OpenDb(Optional ByVal Flags As SQLiteOpenFlags = SQLITE_OPEN_DEFAULT, _
                  Optional ByVal vfsName As String = vbNullString)
    Dim ResultCode As SQLiteResultCodes
    If Flags = SQLITE_OPEN_DEFAULT And Len(vfsName) = 0 Then
        ResultCode = sqlite3_open16(StrPtr(this.DbPathName), this.DbHandle)
    Else
        #If VBA7 Then
            Dim DbPathNamePtr As LongPtr
            Dim vfsNamePtr As LongPtr
        #Else
            Dim DbPathNamePtr As Long
            Dim vfsNamePtr As Long
        #End If
        
        '@Ignore UseMeaningfulName
        Dim DbPathNameUTF8() As Byte
        DbPathNameUTF8 = UTFlib.UTF8BytesFromStr(this.DbPathName)
        DbPathNamePtr = VarPtr(DbPathNameUTF8(0))
        
        If Len(vfsName) = 0 Then
            vfsNamePtr = 0
        Else
            '@Ignore UseMeaningfulName
            Dim vfsNameUTF8() As Byte
            vfsNameUTF8 = UTFlib.UTF8BytesFromStr(vfsName)
            vfsNamePtr = VarPtr(vfsNameUTF8(0))
        End If
        
        ResultCode = sqlite3_open_v2(DbPathNamePtr, this.DbHandle, Flags, vfsNamePtr)
    End If
    If ResultCode <> SQLITE_OK Then ErrInfoRetrieve
    If VERBOUS Then Debug.Print "Open db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
End Sub


Public Sub CloseDb()
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_close(this.DbHandle)
    If ResultCode = SQLITE_OK Then
        this.DbHandle = 0
    Else
        ErrInfoRetrieve
    End If
    If VERBOUS Then Debug.Print "Close db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
End Sub


Public Function ChangesCount(Optional ByVal Total As Boolean = False) As Long
    If Total Then
        ChangesCount = sqlite3_total_changes(this.DbHandle)
    Else
        ChangesCount = sqlite3_changes(this.DbHandle)
    End If
End Function


Public Function ExecuteNonQuery(ByVal SQLQuery As String, _
                       Optional ByRef AffectedRowsCount As Long) As SQLiteResultCodes
    Guard.EmptyString SQLQuery
    
    AffectedRowsCount = ChangesCount(True)
    Dim SQLQueryUTF8() As Byte
    SQLQueryUTF8 = UTFlib.UTF8BytesFromStr(SQLQuery)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_exec(this.DbHandle, VarPtr(SQLQueryUTF8(0)), 0, 0, 0)
    If ResultCode <> SQLITE_OK Then
        ErrInfoRetrieve
        AffectedRowsCount = -1
    Else
        AffectedRowsCount = ChangesCount(True) - AffectedRowsCount
    End If
    ExecuteNonQuery = ResultCode
End Function


Public Function TxnState(Optional ByVal SchemaAlias As String = "main" _
                        ) As SQLiteTxnState
    Dim SchemaUTF8() As Byte
    SchemaUTF8 = UTFlib.UTF8BytesFromStr(SchemaUTF8)
    TxnState = sqlite3_txn_state(this.DbHandle, VarPtr(SchemaUTF8(0)))
End Function


Public Function Begin(Optional ByVal Mode As SQLiteTxnMode = SQLITE_TXN_DEFERRED _
                     ) As SQLiteResultCodes
    Dim TxMode As String
    Select Case Mode
        Case SQLITE_TXN_DEFERRED
            TxMode = vbNullString
        Case SQLITE_TXN_IMMEDIATE
            TxMode = " IMMEDIATE"
        Case SQLITE_TXN_EXCLUSIVE
            TxMode = " EXCLUSIVE"
        Case Else
            Err.Raise ErrNo.InvalidParameterErr, "SQLiteCConnection", _
                "Invalid transaction mode: <" & CStr(Mode) & ">"
    End Select
    Begin = ExecuteNonQuery("BEGIN" & TxMode)
End Function


Public Function Commit() As SQLiteResultCodes
    Commit = ExecuteNonQuery("COMMIT")
End Function


Public Function SavePoint(ByVal SavePointName As String) As SQLiteResultCodes
    Guard.EmptyString SavePointName
    SavePoint = ExecuteNonQuery("SAVEPOINT " & SavePointName)
End Function


Public Function ReleasePoint(ByVal SavePointName As String) As SQLiteResultCodes
    Guard.EmptyString SavePointName
    ReleasePoint = ExecuteNonQuery("RELEASE " & SavePointName)
End Function


Public Function Rollback(Optional ByVal SavePointName As String = vbNullString _
                        ) As SQLiteResultCodes
    Dim SPClause As String
    SPClause = IIf(Len(SavePointName) = 0, vbNullString, " TO " & SavePointName)
    Rollback = ExecuteNonQuery("ROLLBACK" & SPClause)
End Function
