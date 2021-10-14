VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Manages lifecycle of the SQLite Database object (DBO)"
'@Folder "SQLiteC For VBA.Connection"
'@ModuleDescription "Manages lifecycle of the SQLite Database object (DBO)"
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

'''' Custom functions added to SQLite source for testing/verification purposes
#If VBA7 Then
    #If WIN64 Then
    Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As LongLong
    #Else
    Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
    #End If
#Else
Private Declare Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
#End If

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
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
    Statements As Scripting.Dictionary
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
    Set this.Statements = New Scripting.Dictionary
    this.Statements.CompareMode = TextCompare
End Sub


'''' vbNullString is an acceptable StmtName.
'''' Raises an error if StmtName has already been used.
'@Description "Creates a new SQLiteCStatement instance."
Public Function CreateStatement(ByVal StmtName As String) As SQLiteCStatement
Attribute CreateStatement.VB_Description = "Creates a new SQLiteCStatement instance."
    If this.Statements.Exists(StmtName) Then
        Err.Raise KeyAlreadyExistsErr, "SQLiteCConnection", _
                  "Statement name already exists!"
    End If
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = SQLiteCStatement(Me)
    If DbStmt Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCConnection", "Failed to create an SQLiteCStatement instance."
    Set this.Statements(StmtName) = DbStmt
    Set CreateStatement = DbStmt
End Function


'''' vbNullString is an acceptable StmtName (can be used as default).
'@Description "Returns an existing SQLiteCStatement instance or Nothing."
Public Function StmtDb(Optional ByVal StmtName As String = vbNullString _
                      ) As SQLiteCStatement
Attribute StmtDb.VB_Description = "Returns an existing SQLiteCStatement instance or Nothing."
    If this.Statements.Exists(StmtName) Then
        Set StmtDb = this.Statements(StmtName)
    Else
        Set StmtDb = Nothing
    End If
End Function


'@Description "Finalizes all statement objects associated with connection."
Friend Function Finalize() As SQLiteResultCodes
Attribute Finalize.VB_Description = "Finalizes all statement objects associated with connection."
    Dim ResultCode As SQLiteResultCodes
    Dim StmtName As Variant
    Dim StmtDb As SQLiteCStatement
    For Each StmtName In this.Statements.Keys
        Set StmtDb = this.Statements(StmtName)
        ResultCode = StmtDb.Finalize
        If ResultCode <> SQLITE_OK Then ErrInfoRetrieve
    Next StmtName
    Finalize = ResultCode
End Function


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


Public Function ErrInfoRetrieve() As SQLiteResultCodes
    With this.ErrorInfo
        .GetErr
        .PrintErr
        ErrInfoRetrieve = .ErrorCodeEx
    End With
End Function


'''' Reference: https://www.sqlite.org/c3ref/libversion.html
''''
'@Description "Returns SQLite version. No database is necessary."
Public Function Version(Optional ByVal Numeric As Boolean = True) As Variant
Attribute Version.VB_Description = "Returns SQLite version. No database is necessary."
    If Numeric Then
        Version = sqlite3_libversion_number()
    Else
        Version = UTFlib.StrFromUTF8Ptr(sqlite3_libversion())
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/open.html
''''
'@Description "Generates SQLite DBO instance."
Public Sub OpenDb(Optional ByVal Flags As SQLiteOpenFlags = SQLITE_OPEN_DEFAULT, _
                  Optional ByVal vfsName As String = vbNullString)
Attribute OpenDb.VB_Description = "Generates SQLite DBO instance."
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
    '@Ignore FunctionReturnValueDiscarded
    If ResultCode <> SQLITE_OK Then ErrInfoRetrieve
    If VERBOUS Then Debug.Print "Open db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
End Sub


'''' Reference: https://www.sqlite.org/c3ref/close.html
''''
'@Description "Destructs SQLite DBO instance."
Public Sub CloseDb()
Attribute CloseDb.VB_Description = "Destructs SQLite DBO instance."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = Finalize
    Debug.Assert ResultCode = SQLITE_OK
    ResultCode = sqlite3_close(this.DbHandle)
    If ResultCode = SQLITE_OK Then
        this.DbHandle = 0
    Else
        '@Ignore FunctionReturnValueDiscarded
        ErrInfoRetrieve
    End If
    If VERBOUS Then Debug.Print "Close db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
End Sub


'''' Reference: https://www.sqlite.org/c3ref/changes.html
'''' Reference: https://www.sqlite.org/c3ref/total_changes.html
''''
'@Description "Queries the number of changes occured during the last executed statement or since the DBO creation."
Public Function ChangesCount(Optional ByVal Total As Boolean = False) As Long
Attribute ChangesCount.VB_Description = "Queries the number of changes occured during the last executed statement or since the DBO creation."
    If Total Then
        ChangesCount = sqlite3_total_changes(this.DbHandle)
    Else
        ChangesCount = sqlite3_changes(this.DbHandle)
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/exec.html
''''
'@Description "Executes a plain-text non-query via a convenience shortcut."
Public Function ExecuteNonQueryPlain(ByVal SQLQuery As String, _
            Optional ByRef AffectedRows As Long = -1) As SQLiteResultCodes
Attribute ExecuteNonQueryPlain.VB_Description = "Executes a plain-text non-query via a convenience shortcut."
    Guard.EmptyString SQLQuery
    
    If AffectedRows <> -1 Then AffectedRows = ChangesCount(True)
    Dim SQLQueryUTF8() As Byte
    SQLQueryUTF8 = UTFlib.UTF8BytesFromStr(SQLQuery)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_exec(this.DbHandle, VarPtr(SQLQueryUTF8(0)), 0, 0, 0)
    If ResultCode <> SQLITE_OK Then
        '@Ignore FunctionReturnValueDiscarded
        ErrInfoRetrieve
        AffectedRows = -1
    Else
        If AffectedRows <> -1 Then AffectedRows = ChangesCount(True) - AffectedRows
    End If
    ExecuteNonQueryPlain = ResultCode
End Function


'''' Reference: https://www.sqlite.org/c3ref/c_txn_none.html
''''
'@Description "Queries transaction status for the current SQLite DBO."
Public Function TxnState(Optional ByVal SchemaAlias As String = "main" _
                        ) As SQLiteTxnState
Attribute TxnState.VB_Description = "Queries transaction status for the current SQLite DBO."
    '@Ignore UseMeaningfulName
    Dim SchemaUTF8() As Byte
    SchemaUTF8 = UTFlib.UTF8BytesFromStr(SchemaUTF8)
    TxnState = sqlite3_txn_state(this.DbHandle, VarPtr(SchemaUTF8(0)))
End Function


'''' Reference: https://www.sqlite.org/lang_transaction.html
''''
'@Description "Begins SQLite transaction."
Public Function Begin(Optional ByVal Mode As SQLiteTxnMode = SQLITE_TXN_DEFERRED _
                     ) As SQLiteResultCodes
Attribute Begin.VB_Description = "Begins SQLite transaction."
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
    Begin = ExecuteNonQueryPlain("BEGIN" & TxMode)
End Function


'''' Reference: https://www.sqlite.org/lang_transaction.html
''''
'@Description "Commits all SQLite DBO transactions."
Public Function Commit() As SQLiteResultCodes
Attribute Commit.VB_Description = "Commits all SQLite DBO transactions."
    Commit = ExecuteNonQueryPlain("COMMIT")
End Function


'''' Reference: https://www.sqlite.org/lang_savepoint.html
''''
'@Description "Begins a named and nestable SQLite transaction."
Public Function SavePoint(ByVal SavePointName As String) As SQLiteResultCodes
Attribute SavePoint.VB_Description = "Begins a named and nestable SQLite transaction."
    Guard.EmptyString SavePointName
    SavePoint = ExecuteNonQueryPlain("SAVEPOINT " & SavePointName)
End Function


'''' Reference: https://www.sqlite.org/lang_savepoint.html
''''
'@Description "Removes rollback savepoints from the transaction stack up to the first point with matching name."
Public Function ReleasePoint(ByVal SavePointName As String) As SQLiteResultCodes
Attribute ReleasePoint.VB_Description = "Removes rollback savepoints from the transaction stack up to the first point with matching name."
    Guard.EmptyString SavePointName
    ReleasePoint = ExecuteNonQueryPlain("RELEASE " & SavePointName)
End Function


'''' Reference: https://www.sqlite.org/lang_transaction.html
'''' Reference: https://www.sqlite.org/lang_savepoint.html
''''
'@Description "Rollbacks savepoints from the transaction stack up to the first point with matching name or the entire transaction."
Public Function Rollback(Optional ByVal SavePointName As String = vbNullString _
                        ) As SQLiteResultCodes
Attribute Rollback.VB_Description = "Rollbacks savepoints from the transaction stack up to the first point with matching name or the entire transaction."
    Dim SPClause As String
    SPClause = IIf(Len(SavePointName) = 0, vbNullString, " TO " & SavePointName)
    Rollback = ExecuteNonQueryPlain("ROLLBACK" & SPClause)
End Function


'''' This routine uses a non-standard export added to a the SQLite amalgamation
'''' yielding a custom-built library. It returns SQLite version as a 64-bit int.
'''' It can be used for testing purposes, especially for the x32 version.
''''
'@Ignore UseMeaningfulName
Public Function VersionI64() As Variant
    Dim VerNum As Variant
    VerNum = sqlite3_libversion_number_i64()
    If VarType(VerNum) = vbCurrency Then VerNum = VerNum * 10000
    VersionI64 = VerNum
End Function
