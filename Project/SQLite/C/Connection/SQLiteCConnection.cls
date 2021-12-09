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
'@Folder "SQLite.C.Connection"
'@ModuleDescription "Manages lifecycle of the SQLite Database object (DBO)"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, _
    ByVal iFlags As Long, ByVal zVfs As LongPtr) As SQLiteResultCodes ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_db_filename Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSchema As LongPtr) As LongPtr
Private Declare PtrSafe Function sqlite3_db_readonly Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSchema As LongPtr) As SQLiteDbAccess
'''' Transactions
Private Declare PtrSafe Function sqlite3_txn_state Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSchema As LongPtr) As SQLiteTxnState
'''' Statements
Private Declare PtrSafe Function sqlite3_exec Lib "SQLite3" (ByVal hDb As LongPtr, ByVal zSql As LongPtr, _
    ByVal Callback As LongPtr, ByVal pArg As LongPtr, ByRef pzErrMsg As LongPtr) As SQLiteResultCodes
'''' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
'''' Error information
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteResultCodes) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
#Else
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As SQLiteResultCodes ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, _
    ByVal zVfs As Long) As SQLiteResultCodes ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_db_filename Lib "SQLite3" (ByVal hDb As Long, ByVal zSchema As Long) As Long
Private Declare Function sqlite3_db_readonly Lib "SQLite3" (ByVal hDb As Long, ByVal zSchema As Long) As SQLiteDbAccess
'''' Transactions
Private Declare Function sqlite3_txn_state Lib "SQLite3" (ByVal hDb As Long, ByVal zSchema As Long) As SQLiteTxnState
'''' Statements
Private Declare Function sqlite3_exec Lib "SQLite3" (ByVal hDb As Long, ByVal zSql As Long, _
    ByVal Callback As Long, ByVal pArg As Long, ByRef pzErrMsg As Long) As SQLiteResultCodes
'''' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As Long) As Long
'''' Error information
Private Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteResultCodes) As Long ' PtrUtf8String
Private Declare Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
#End If

Public Enum SQLiteOpenFlags
    SQLITE_OPEN_READONLY = &H1&
    SQLITE_OPEN_READWRITE = &H2&
    SQLITE_OPEN_CREATE = &H4&
    SQLITE_OPEN_URI = &H40&
    SQLITE_OPEN_MEMORY = &H80&
    SQLITE_OPEN_NOMUTEX = &H8000&
    SQLITE_OPEN_FULLMUTEX = &H10000
    SQLITE_OPEN_SHAREDCACHE = &H20000
    SQLITE_OPEN_PRIVATECACHE = &H40000
    SQLITE_OPEN_NOFOLLOW = &H1000000
    SQLITE_OPEN_DEFAULT = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
End Enum

Public Enum SQLiteTxnType
    SQLITE_TXN_DEFERRED = 0&
    SQLITE_TXN_IMMEDIATE = 1&
    SQLITE_TXN_EXCLUSIVE = 2&
End Enum

Public Enum SQLiteTxnState
    SQLITE_TXN_NULL = -1&
    SQLITE_TXN_NONE = 0&
    SQLITE_TXN_READ = 1&
    SQLITE_TXN_WRITE = 2&
End Enum

Public Enum SQLiteDbAccess
    SQLITE_DB_NULL = -2
    SQLITE_DB_NONE = -1
    SQLITE_DB_FULL = 0
    SQLITE_DB_READ = 1
End Enum

Public Enum SQLiteJournalMode
    SQLITE_PAGER_JOURNALMODE_DELETE = 0     '/* Commit by deleting journal file */
    SQLITE_PAGER_JOURNALMODE_PERSIST = 1    '/* Commit by zeroing journal header */
    SQLITE_PAGER_JOURNALMODE_OFF = 2        '/* Journal omitted.  */
    SQLITE_PAGER_JOURNALMODE_TRUNCATE = 3   '/* Commit by truncating journal */
    SQLITE_PAGER_JOURNALMODE_MEMORY = 4     '/* In-memory journal file */
    SQLITE_PAGER_JOURNALMODE_WAL = 5        '/* Use write-ahead logging */
End Enum

Private Type TSQLiteCConnection
    DbPathName As String                    ' Resolved file pathname of the "main" database.
    #If VBA7 Then
        DbHandle As LongPtr
    #Else
        DbHandle As Long
    #End If
    Statements As Scripting.Dictionary
    ErrInfo As SQLiteCErr
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
    this.DbPathName = DbPathName
    Set this.Statements = New Scripting.Dictionary
    this.Statements.CompareMode = TextCompare
End Sub

Friend Sub CleanUp()
    With this
        Dim DbStmt As SQLiteCStatement
        Dim StmtName As Variant
        For Each StmtName In .Statements.Keys
            Set DbStmt = .Statements(StmtName)
            DbStmt.CleanUp
        Next StmtName
        Set DbStmt = Nothing
        .Statements.RemoveAll
        Set .Statements = Nothing
        Set .Statements = New Scripting.Dictionary
        .Statements.CompareMode = TextCompare
    End With
    Dim ResultCode As SQLiteResultCodes
    ResultCode = CloseDb
    Debug.Assert ResultCode = SQLITE_OK
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

'''' Reference: https://www.sqlite.org/c3ref/open.html
'''' Reference: https://www.sqlite.org/vfs.html
''''
'@Description "Generates SQLite DBO instance."
Public Function OpenDb( _
            Optional ByVal Flags As SQLiteOpenFlags = SQLITE_OPEN_DEFAULT, _
            Optional ByVal vfsName As String = vbNullString _
            ) As SQLiteResultCodes
Attribute OpenDb.VB_Description = "Generates SQLite DBO instance."
    If this.DbHandle > 0 Then
        OpenDb = SQLITE_OK
        Exit Function
    End If
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
        Dim DbPathNameUTF8B() As Byte
        DbPathNameUTF8B = UTFlib.UTF8BytesFromStr(this.DbPathName)
        DbPathNamePtr = VarPtr(DbPathNameUTF8B(0))
        If Len(vfsName) = 0 Then
            vfsNamePtr = 0
        Else
            Dim vfsNameUTF8B() As Byte
            vfsNameUTF8B = UTFlib.UTF8BytesFromStr(vfsName)
            vfsNamePtr = VarPtr(vfsNameUTF8B(0))
        End If
        ResultCode = sqlite3_open_v2(DbPathNamePtr, this.DbHandle, Flags, vfsNamePtr)
    End If
    If ResultCode <> SQLITE_OK Then ResultCode = ErrInfoRetrieve
    Dim LogMsg As String
    LogMsg = "Open db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
    Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    OpenDb = ResultCode
End Function

'''' Reference: https://www.sqlite.org/c3ref/close.html
''''
'@Description "Destructs SQLite DBO instance."
Public Function CloseDb() As SQLiteResultCodes
Attribute CloseDb.VB_Description = "Destructs SQLite DBO instance."
    If this.DbHandle = 0 Then
        CloseDb = SQLITE_OK
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
    ResultCode = Finalize
    Debug.Assert ResultCode = SQLITE_OK
    If this.DbHandle <> 0 Then
        ResultCode = sqlite3_close(this.DbHandle)
        If ResultCode = SQLITE_OK Then
            this.DbHandle = 0
        Else
            ResultCode = ErrInfoRetrieve
        End If
    End If
    Dim LogMsg As String
    LogMsg = "Close db status: <" & SQLiteCRC.CodeToName(ResultCode) & ">"
    Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    CloseDb = ResultCode
End Function

Public Property Get ErrInfo() As SQLiteCErr
    ErrInfo = this.ErrInfo
End Property

'''' Reference: https://www.sqlite.org/c3ref/errcode.html
''''
Public Sub GetErr()
    FixGuard.DbNotOpened Me, "SQLiteCConnection/GetErr"
    With this.ErrInfo
        .ErrorCode = sqlite3_errcode(this.DbHandle)
        .ErrorCodeName = SQLiteCRC.CodeToName(.ErrorCode)
        .ErrorCodeEx = sqlite3_extended_errcode(this.DbHandle)
        .ErrorCodeExName = SQLiteCRC.CodeToName(.ErrorCodeEx)
        .ErrorName = .ErrorCodeExName
        .ErrorMessage = UTFlib.StrFromUTF8Ptr(sqlite3_errmsg(this.DbHandle))
        .ErrorString = UTFlib.StrFromUTF8Ptr(sqlite3_errstr(.ErrorCode))
    End With
End Sub

Public Sub PrintErr()
    With this.ErrInfo
        Logger.Logg "========== Error Details ========== " & vbNewLine & _
                    "ErrCode:          " & CStr(.ErrorCode) & vbNewLine & _
                    "ErrCode Extended: " & CStr(.ErrorCodeEx) & vbNewLine & _
                    "Error Name:       " & .ErrorName & vbNewLine & _
                    "ErrStr:           " & .ErrorString & vbNewLine & _
                    "ErrMsg:           " & .ErrorMessage & vbNewLine _
                    , , DEBUGLEVEL_ERROR
    End With
End Sub

Public Function ErrInfoRetrieve(Optional ByVal PrintDebug As Boolean = True _
                               ) As SQLiteResultCodes
    GetErr
    If PrintDebug Then PrintErr
    ErrInfoRetrieve = this.ErrInfo.ErrorCodeEx
End Function

'''' Reference: https://www.sqlite.org/c3ref/libversion.html
''''
'@Description "Returns SQLite version. No database is necessary."
Public Function VersionNumber() As Long
Attribute VersionNumber.VB_Description = "Returns SQLite version. No database is necessary."
    VersionNumber = sqlite3_libversion_number()
End Function

'@Description "Retrieves database path for SchemaAlias."
Public Function AttachedDbPathName(Optional ByVal SchemaAlias As String = "main" _
                                  ) As String
Attribute AttachedDbPathName.VB_Description = "Retrieves database path for SchemaAlias."
    FixGuard.DbNotOpened Me, "SQLiteCConnection/AttachedDbPathName"
    Dim SchemaUTF8() As Byte
    SchemaUTF8 = UTFlib.UTF8BytesFromStr(SchemaAlias)
    AttachedDbPathName = UTFlib.StrFromUTF8Ptr( _
        sqlite3_db_filename(this.DbHandle, VarPtr(SchemaUTF8(0))))
End Function

'''' Reference: https://www.sqlite.org/c3ref/db_readonly.html
''''
'@Description "Queries database access mode."
Public Function AccessMode(Optional ByVal SchemaAlias As String = "main" _
                          ) As SQLiteDbAccess
Attribute AccessMode.VB_Description = "Queries database access mode."
    FixGuard.DbNotOpened Me, "SQLiteCConnection/AccessMode"
    Dim SchemaUTF8() As Byte
    SchemaUTF8 = UTFlib.UTF8BytesFromStr(SchemaAlias)
    AccessMode = sqlite3_db_readonly(this.DbHandle, VarPtr(SchemaUTF8(0)))
End Function

'''' Reference: https://www.sqlite.org/c3ref/changes.html
'''' Reference: https://www.sqlite.org/c3ref/total_changes.html
''''
'@Description "Queries the number of changes occured during the last executed statement or since the DBO creation."
Public Function ChangesCount(Optional ByVal Total As Boolean = False) As Long
Attribute ChangesCount.VB_Description = "Queries the number of changes occured during the last executed statement or since the DBO creation."
    FixGuard.DbNotOpened Me, "SQLiteCConnection/ChangesCount"
    If Total Then
        ChangesCount = sqlite3_total_changes(this.DbHandle)
    Else
        ChangesCount = sqlite3_changes(this.DbHandle)
    End If
End Function

'''' This interface does not support parametrized queries as opposed to the
'''' similar interface on the Statemnt object. This is because this interface
'''' uses a "shortcut" API not supporting parameterized quiries.
''''
'''' This interface, however, supports execution of multi-statement SQL commands.
'''' This difference is also due to the differences between the underlying APIs.
''''
'''' Reference: https://www.sqlite.org/c3ref/exec.html
''''
'@Description "Executes a plain-text non-query via a convenience shortcut."
Public Function ExecuteNonQueryPlain(ByVal SQLQuery As String, _
            Optional ByRef AffectedRows As Long = -1) As SQLiteResultCodes
Attribute ExecuteNonQueryPlain.VB_Description = "Executes a plain-text non-query via a convenience shortcut."
    Guard.ExpressionErr Len(SQLQuery) > 0, InvalidParameterErr, _
        "SQLiteCConnection/ExecuteNonQueryPlain", "Query string is empty."
    FixGuard.DbNotOpened Me, "SQLiteCConnection/ExecuteNonQueryPlain"
    If AffectedRows <> -1 Then AffectedRows = ChangesCount(True)
    Dim SQLQueryUTF8() As Byte
    SQLQueryUTF8 = UTFlib.UTF8BytesFromStr(SQLQuery)
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_exec(this.DbHandle, VarPtr(SQLQueryUTF8(0)), 0, 0, 0)
    If ResultCode <> SQLITE_OK Then
        ResultCode = ErrInfoRetrieve
        AffectedRows = -1
    Else
        If AffectedRows <> -1 Then AffectedRows = ChangesCount(True) - AffectedRows
    End If
    ExecuteNonQueryPlain = ResultCode
End Function

'@Description "Attaches database to the current connection."
Public Function Attach(ByVal DbPathName As String, _
                      Optional ByVal SchemaAlias As String = vbNullString _
                      ) As SQLiteResultCodes
Attribute Attach.VB_Description = "Attaches database to the current connection."
    Dim FilePathName As String
    Dim PathCheck As LiteFSCheck
    Set PathCheck = LiteFSCheck(DbPathName, True)
    Guard.ExpressionTrueErr PathCheck.ErrNumber <> 0, ErrNo.UnknownClassErr, _
        "SQLiteCConnection/Attach", "Failed to resolve db path."
    FilePathName = PathCheck.DatabasePathName
    Dim SQLQuery As String
    SQLQuery = SQLlib.Attach(FilePathName, SchemaAlias)
    Attach = ExecuteNonQueryPlain(SQLQuery)
End Function

'@Description "Detaches a database from the current connection."
Public Function Detach(ByVal SchemaAlias As String) As SQLiteResultCodes
Attribute Detach.VB_Description = "Detaches a database from the current connection."
    Guard.EmptyString SchemaAlias
    Dim SQLQuery As String
    SQLQuery = SQLlib.Detach(SchemaAlias)
    Detach = ExecuteNonQueryPlain(SQLQuery)
End Function

'@Description "Vacuums a database in-place or into a new file."
Public Function Vacuum( _
            Optional ByVal SrcSchemaAlias As String = vbNullString, _
            Optional ByVal DstDbPathName As String = vbNullString _
            ) As SQLiteResultCodes
Attribute Vacuum.VB_Description = "Vacuums a database in-place or into a new file."
    Dim FilePathName As String
    If Len(DstDbPathName) > 0 Then
        Dim PathCheck As LiteFSCheck
        Set PathCheck = LiteFSCheck(DstDbPathName, True)
        Guard.ExpressionTrueErr PathCheck.ErrNumber <> 0, ErrNo.UnknownClassErr, _
            "SQLiteCConnection/Vacuum", "Failed to resolve db path."
        FilePathName = PathCheck.DatabasePathName
        '@Ignore SelfAssignedDeclaration
        Dim fso As New Scripting.FileSystemObject
        If fso.FileExists(FilePathName) Then
            Guard.ExpressionTrueErr fso.GetFile(FilePathName).Size > 0, _
                ErrNo.FileAlreadyExistsErr, "SQLiteCConnection/Vacuum", _
                "Non-empty file already exists."
        End If
    End If
    Dim SQLQuery As String
    SQLQuery = SQLlib.Vacuum(SrcSchemaAlias, FilePathName)
    Vacuum = ExecuteNonQueryPlain(SQLQuery)
End Function

'''' vbNullString is an acceptable StmtName.
'''' Raises an error if StmtName has already been used.
'@Description "Creates a new SQLiteCStatement instance."
Public Function CreateStatement(ByVal StmtName As String) As SQLiteCStatement
Attribute CreateStatement.VB_Description = "Creates a new SQLiteCStatement instance."
    Guard.ExpressionTrueErr this.Statements.Exists(StmtName), _
        ErrNo.KeyAlreadyExistsErr, "SQLiteCConnection/CreateStatement", _
        "Statement name already exists!"
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = SQLiteCStatement(Me)
    Guard.NullReference DbStmt, "SQLiteCConnection/CreateStatement", _
        "Failed to create an SQLiteCStatement instance."
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

'@Description "Shortcut returning ILiteADO interface instance."
Public Property Get ExecADO() As ILiteADO
Attribute ExecADO.VB_Description = "Shortcut returning ILiteADO interface instance."
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = StmtDb("ILiteADO")
    If DbStmt Is Nothing Then Set DbStmt = CreateStatement("ILiteADO")
    Set ExecADO = DbStmt
End Property

'@Description "Finalizes all statement objects associated with connection."
Friend Function Finalize() As SQLiteResultCodes
Attribute Finalize.VB_Description = "Finalizes all statement objects associated with connection."
    If this.Statements Is Nothing Then
        Finalize = SQLITE_OK
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
    Dim StmtName As Variant
    Dim DbStmt As SQLiteCStatement
    For Each StmtName In this.Statements.Keys
        Set DbStmt = this.Statements(StmtName)
        ResultCode = DbStmt.Finalize
        If ResultCode <> SQLITE_OK Then ResultCode = ErrInfoRetrieve
    Next StmtName
    ResultCode = SQLiteCStatement.PreparedStatementsGet(Me, True)
    Finalize = ResultCode
End Function

'@Description "Sets SQLite journal mode."
Public Function JournalModeSet(ByVal JournalModeCode As SQLiteJournalMode, _
                      Optional ByVal SchemaAlias As String = "main" _
                      ) As SQLiteResultCodes
Attribute JournalModeSet.VB_Description = "Sets SQLite journal mode."
    Dim DbSchemaPrefix As String
    If Len(SchemaAlias) > 0 And SchemaAlias <> "main" Then
        DbSchemaPrefix = SchemaAlias & "."
    Else
        DbSchemaPrefix = vbNullString
    End If
    Dim JournalMode As String
    JournalMode = Array("DELETE", "PERSIST", "OFF", _
                        "TRUNCATE", "MEMORY", "WAL")(JournalModeCode)
    Dim SQLQuery As String
    SQLQuery = "PRAGMA " & DbSchemaPrefix & "journal_mode = '" & JournalMode & "'"
    JournalModeSet = ExecuteNonQueryPlain(SQLQuery)
End Function

'''' Reference: https://www.sqlite.org/c3ref/c_txn_none.html
''''
'@Description "Queries transaction status for the current SQLite DBO."
Public Function TxnState(Optional ByVal SchemaAlias As String = "main" _
                        ) As SQLiteTxnState
Attribute TxnState.VB_Description = "Queries transaction status for the current SQLite DBO."
    Dim SchemaUTF8() As Byte
    SchemaUTF8 = UTFlib.UTF8BytesFromStr(SchemaAlias)
    TxnState = sqlite3_txn_state(this.DbHandle, VarPtr(SchemaUTF8(0)))
End Function

'''' Reference: https://www.sqlite.org/lang_transaction.html
''''
'@Description "Begins SQLite transaction."
Public Function Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED _
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
    SPClause = IIf(Len(SavePointName) = 0, vbNullString, " TO SAVEPOINT " & SavePointName)
    Rollback = ExecuteNonQueryPlain("ROLLBACK" & SPClause)
End Function

Public Function DbIsLocked() As Variant
    If TxnState("main") <> SQLITE_TXN_NONE Then
        DbIsLocked = CVErr(ErrNo.AdoInTransactionErr)
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
    ResultCode = Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode = SQLITE_BUSY Then
        DbIsLocked = True
    ElseIf ResultCode = SQLITE_OK Then
        If Commit = SQLITE_OK Then
            DbIsLocked = False
        Else
            DbIsLocked = CVErr(this.ErrInfo.ErrorCode)
        End If
    Else
        DbIsLocked = CVErr(this.ErrInfo.ErrorCode)
    End If
End Function
