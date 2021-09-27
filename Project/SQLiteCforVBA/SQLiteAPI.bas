Attribute VB_Name = "SQLiteAPI"
'@Folder "SQLiteCforVBA"
Option Explicit


Public Const SQLITE_STATIC      As Long = 0
Public Const SQLITE_TRANSIENT   As Long = -1

Public Enum SQLiteErrors
    SQLITE_OK = 0&
    SQLITE_ERROR = 1&
    SQLITE_INTERNAL = 2&
    SQLITE_PERM = 3&
    SQLITE_ABORT = 4&
    SQLITE_BUSY = 5&
    SQLITE_LOCKED = 6&
    SQLITE_NOMEM = 7&
    SQLITE_READONLY = 8&
    SQLITE_INTERRUPT = 9&
    SQLITE_IOERR = 10&
    SQLITE_CORRUPT = 11&
    SQLITE_NOTFOUND = 12&
    SQLITE_FULL = 13&
    SQLITE_CANTOPEN = 14&
    SQLITE_PROTOCOL = 15&
    SQLITE_EMPTY = 16&
    SQLITE_SCHEMA = 17&
    SQLITE_TOOBIG = 18&
    SQLITE_CONSTRAINT = 19&
    SQLITE_MISMATCH = 20&
    SQLITE_MISUSE = 21&
    SQLITE_NOLFS = 22&
    SQLITE_AUTH = 23&
    SQLITE_FORMAT = 24&
    SQLITE_RANGE = 25&
    SQLITE_NOTADB = 26&
    SQLITE_NOTICE = 27&
    SQLITE_WARNING = 28&
    SQLITE_ROW = 100&
    SQLITE_DONE = 101&
    SQLITE_ERROR_MISSING_COLLSEQ = SQLITE_ERROR + 1 * 256
    SQLITE_ERROR_RETRY = SQLITE_ERROR + 2 * 256
    SQLITE_ERROR_SNAPSHOT = SQLITE_ERROR + 3 * 256
    SQLITE_IOERR_READ = SQLITE_IOERR + 1 * 256
    SQLITE_IOERR_SHORT_READ = SQLITE_IOERR + 2 * 256
    SQLITE_IOERR_WRITE = SQLITE_IOERR + 3 * 256
    SQLITE_IOERR_FSYNC = SQLITE_IOERR + 4 * 256
    SQLITE_IOERR_DIR_FSYNC = SQLITE_IOERR + 5 * 256
    SQLITE_IOERR_TRUNCATE = SQLITE_IOERR + 6 * 256
    SQLITE_IOERR_FSTAT = SQLITE_IOERR + 7 * 256
    SQLITE_IOERR_UNLOCK = SQLITE_IOERR + 8 * 256
    SQLITE_IOERR_RDLOCK = SQLITE_IOERR + 9 * 256
    SQLITE_IOERR_DELETE = SQLITE_IOERR + 10 * 256
    SQLITE_IOERR_BLOCKED = SQLITE_IOERR + 11 * 256
    SQLITE_IOERR_NOMEM = SQLITE_IOERR + 12 * 256
    SQLITE_IOERR_ACCESS = SQLITE_IOERR + 13 * 256
    SQLITE_IOERR_CHECKRESERVEDLOCK = SQLITE_IOERR + 14 * 256
    SQLITE_IOERR_LOCK = SQLITE_IOERR + 15 * 256
    SQLITE_IOERR_CLOSE = SQLITE_IOERR + 16 * 256
    SQLITE_IOERR_DIR_CLOSE = SQLITE_IOERR + 17 * 256
    SQLITE_IOERR_SHMOPEN = SQLITE_IOERR + 18 * 256
    SQLITE_IOERR_SHMSIZE = SQLITE_IOERR + 19 * 256
    SQLITE_IOERR_SHMLOCK = SQLITE_IOERR + 20 * 256
    SQLITE_IOERR_SHMMAP = SQLITE_IOERR + 21 * 256
    SQLITE_IOERR_SEEK = SQLITE_IOERR + 22 * 256
    SQLITE_IOERR_DELETE_NOENT = SQLITE_IOERR + 23 * 256
    SQLITE_IOERR_MMAP = SQLITE_IOERR + 24 * 256
    SQLITE_IOERR_GETTEMPPATH = SQLITE_IOERR + 25 * 256
    SQLITE_IOERR_CONVPATH = SQLITE_IOERR + 26 * 256
    SQLITE_IOERR_VNODE = SQLITE_IOERR + 27 * 256
    SQLITE_IOERR_AUTH = SQLITE_IOERR + 28 * 256
    SQLITE_IOERR_BEGIN_ATOMIC = SQLITE_IOERR + 29 * 256
    SQLITE_IOERR_COMMIT_ATOMIC = SQLITE_IOERR + 30 * 256
    SQLITE_IOERR_ROLLBACK_ATOMIC = SQLITE_IOERR + 31 * 256
    SQLITE_IOERR_DATA = SQLITE_IOERR + 32 * 256
    SQLITE_IOERR_CORRUPTFS = SQLITE_IOERR + 33 * 256
    SQLITE_LOCKED_SHAREDCACHE = SQLITE_LOCKED + 1 * 256
    SQLITE_LOCKED_VTAB = SQLITE_LOCKED + 2 * 256
    SQLITE_BUSY_RECOVERY = SQLITE_BUSY + 1 * 256
    SQLITE_BUSY_SNAPSHOT = SQLITE_BUSY + 2 * 256
    SQLITE_BUSY_TIMEOUT = SQLITE_BUSY + 3 * 256
    SQLITE_CANTOPEN_NOTEMPDIR = SQLITE_CANTOPEN + 1 * 256
    SQLITE_CANTOPEN_ISDIR = SQLITE_CANTOPEN + 2 * 256
    SQLITE_CANTOPEN_FULLPATH = SQLITE_CANTOPEN + 3 * 256
    SQLITE_CANTOPEN_CONVPATH = SQLITE_CANTOPEN + 4 * 256
    SQLITE_CANTOPEN_DIRTYWAL = SQLITE_CANTOPEN + 5 * 256
    SQLITE_CANTOPEN_SYMLINK = SQLITE_CANTOPEN + 6 * 256
    SQLITE_CORRUPT_VTAB = SQLITE_CORRUPT + 1 * 256
    SQLITE_CORRUPT_SEQUENCE = SQLITE_CORRUPT + 2 * 256
    SQLITE_CORRUPT_INDEX = SQLITE_CORRUPT + 3 * 256
    SQLITE_READONLY_RECOVERY = SQLITE_READONLY + 1 * 256
    SQLITE_READONLY_CANTLOCK = SQLITE_READONLY + 2 * 256
    SQLITE_READONLY_ROLLBACK = SQLITE_READONLY + 3 * 256
    SQLITE_READONLY_DBMOVED = SQLITE_READONLY + 4 * 256
    SQLITE_READONLY_CANTINIT = SQLITE_READONLY + 5 * 256
    SQLITE_READONLY_DIRECTORY = SQLITE_READONLY + 6 * 256
    SQLITE_ABORT_ROLLBACK = SQLITE_ABORT + 2 * 256
    SQLITE_CONSTRAINT_CHECK = SQLITE_CONSTRAINT + 1 * 256
    SQLITE_CONSTRAINT_COMMITHOOK = SQLITE_CONSTRAINT + 2 * 256
    SQLITE_CONSTRAINT_FOREIGNKEY = SQLITE_CONSTRAINT + 3 * 256
    SQLITE_CONSTRAINT_FUNCTION = SQLITE_CONSTRAINT + 4 * 256
    SQLITE_CONSTRAINT_NOTNULL = SQLITE_CONSTRAINT + 5 * 256
    SQLITE_CONSTRAINT_PRIMARYKEY = SQLITE_CONSTRAINT + 6 * 256
    SQLITE_CONSTRAINT_TRIGGER = SQLITE_CONSTRAINT + 7 * 256
    SQLITE_CONSTRAINT_UNIQUE = SQLITE_CONSTRAINT + 8 * 256
    SQLITE_CONSTRAINT_VTAB = SQLITE_CONSTRAINT + 9 * 256
    SQLITE_CONSTRAINT_ROWID = SQLITE_CONSTRAINT + 10 * 256
    SQLITE_CONSTRAINT_PINNED = SQLITE_CONSTRAINT + 11 * 256
    SQLITE_NOTICE_RECOVER_WAL = SQLITE_NOTICE + 1 * 256
    SQLITE_NOTICE_RECOVER_ROLLBACK = SQLITE_NOTICE + 2 * 256
    SQLITE_WARNING_AUTOINDEX = SQLITE_WARNING + 1 * 256
    SQLITE_AUTH_USER = SQLITE_AUTH + 1 * 256
    SQLITE_OK_LOAD_PERMANENTLY = SQLITE_OK + 1 * 256
    SQLITE_OK_SYMLINK = SQLITE_OK + 2 * 256
End Enum


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
End Enum


Public Enum SQLiteTypes
    SQLITE_INTEGER = 1&
    SQLITE_FLOAT = 2&
    SQLITE_TEXT = 3&
    SQLITE_BLOB = 4&
    SQLITE_NULL = 5&
End Enum


#If VBA7 Then

' SQLite library version
Public Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String

' Database connections
Public Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Public Declare PtrSafe Function DbClose Lib "SQLite3" Alias "sqlite3_close" (ByVal hDb As LongPtr) As Long

' Database connection error info
Public Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Public Declare PtrSafe Function sqlite3_errstr Lib "SQLite3" (ByVal ErrorCode As Long) As LongPtr ' PtrUtf8String
Public Declare PtrSafe Function ErrCode Lib "SQLite3" Alias "sqlite3_errcode" (ByVal hDb As LongPtr) As Long
Public Declare PtrSafe Function ErrCodeExtended Lib "SQLite3" Alias "sqlite3_extended_errcode" (ByVal hDb As LongPtr) As Long

' Database connection change counts
Public Declare PtrSafe Function Changes Lib "SQLite3" Alias "sqlite3_changes" (ByVal hDb As LongPtr) As Long
Public Declare PtrSafe Function ChangesTotal Lib "SQLite3" Alias "sqlite3_total_changes" (ByVal hDb As LongPtr) As Long

' Statements
Public Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
Public Declare PtrSafe Function StmtStep Lib "SQLite3" Alias "sqlite3_step" (ByVal hStmt As LongPtr) As Long
Public Declare PtrSafe Function StmtReset Lib "SQLite3" Alias "sqlite3_reset" (ByVal hStmt As LongPtr) As Long
Public Declare PtrSafe Function StmtFinalize Lib "SQLite3" Alias "sqlite3_finalize" (ByVal hStmt As LongPtr) As Long

' Statement column access (0-based indices)
Public Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString

Public Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Public Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function ColumnDouble Lib "SQLite3" Alias "sqlite3_column_double" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Public Declare PtrSafe Function ColumnInt32 Lib "SQLite3" Alias "sqlite3_column_int" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Public Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Public Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Public Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
Public Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
Public Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Public Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Double) As Long
Public Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Long) As Long
Public Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As LongLong) As Long
Public Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

'Backup
Public Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Public Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Public Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Public Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long

#Else

' SQLite library version
Public Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String

' Database connections
Public Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Public Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Public Declare Function DbClose Lib "SQLite3" Alias "sqlite3_close" (ByVal hDb As Long) As Long

' Database connection error info
Public Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
Public Declare Function sqlite3_errstr Lib "SQLite3" (ByVal ErrorCode As Long) As Long ' PtrUtf8String
Public Declare Function ErrCode Lib "SQLite3" Alias "sqlite3_errcode" (ByVal hDb As Long) As Long
Public Declare Function ErrCodeExtended Lib "SQLite3" Alias "sqlite3_extended_errcode" (ByVal hDb As Long) As Long

' Database connection change counts
Public Declare Function Changes Lib "SQLite3" Alias "sqlite3_changes" (ByVal hDb As Long) As Long
Public Declare Function ChangesTotal Lib "SQLite3" Alias "sqlite3_total_changes" (ByVal hDb As Long) As Long

' Statements
Public Declare Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
Public Declare Function StmtStep Lib "SQLite3" Alias "sqlite3_step" (ByVal hStmt As Long) As Long
Public Declare Function StmtReset Lib "SQLite3" Alias "sqlite3_reset" (ByVal hStmt As Long) As Long
Public Declare Function StmtFinalize Lib "SQLite3" Alias "sqlite3_finalize" (ByVal hStmt As Long) As Long

' Statement column access (0-based indices)
Public Declare Function ColumnCount Lib "SQLite3" Alias "sqlite3_column_count" (ByVal hStmt As Long) As Long
Public Declare Function ColumnType Lib "SQLite3" Alias "sqlite3_column_type" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString

Public Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Public Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Public Declare Function ColumnDouble Lib "SQLite3" Alias "sqlite3_column_double" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Public Declare Function ColumnInt32 Lib "SQLite3" Alias "sqlite3_column_int" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
Public Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Public Declare Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Public Declare Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Public Declare Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Public Declare Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As Long, ByVal paramName As Long) As Long
Public Declare Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Public Declare Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Public Declare Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Public Declare Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Double) As Long
Public Declare Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Long) As Long
Public Declare Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Currency) As Long
Public Declare Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Public Declare Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Public Declare Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
Public Declare Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As Long) As Long

'Backup
Public Declare Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Public Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Public Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Public Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
Public Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
Public Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long

#End If


'''' =============================================================
'''' =========================== Meta ============================
'''' =============================================================

Public Function LibVersion() As String
    LibVersion = UTF8PtrToString(sqlite3_libversion())
End Function

'''' =============================================================
'''' ==================== Database connection ====================
'''' =============================================================

#If VBA7 Then
Public Function DbOpen16(ByVal FileName As String, ByRef dbHandle As LongPtr) As Long
#Else
Public Function DbOpen16(ByVal FileName As String, ByRef dbHandle As Long) As Long
#End If
    DbOpen16 = sqlite3_open16(StrPtr(FileName), dbHandle)
End Function

#If VBA7 Then
Public Function DbOpenV2(ByVal FileName As String, ByRef dbHandle As LongPtr, ByVal Flags As Long, ByVal vfsName As String) As Long
#Else
Public Function DbOpenV2(ByVal FileName As String, ByRef dbHandle As Long, ByVal Flags As Long, ByVal vfsName As String) As Long
#End If
    Dim FileNameBytes() As Byte
    Dim vfsNameBytes() As Byte
    FileNameBytes = StringToUtf8Bytes(FileName)
    If Len(vfsName) = 0 Then
        DbOpenV2 = sqlite3_open_v2(VarPtr(FileNameBytes(0)), dbHandle, Flags, 0)
    Else
        vfsNameBytes = StringToUtf8Bytes(vfsName)
        DbOpenV2 = sqlite3_open_v2(VarPtr(FileNameBytes(0)), dbHandle, Flags, VarPtr(vfsNameBytes(0)))
    End If
End Function

'''' =============================================================
'''' ===================== Error Information =====================
'''' =============================================================

#If VBA7 Then
Public Function ErrMsg(ByVal dbHandle As LongPtr) As String
#Else
Public Function ErrMsg(ByVal dbHandle As Long) As String
#End If
    ErrMsg = UTF8PtrToString(sqlite3_errmsg(dbHandle))
End Function

Public Function ErrStr(ByVal ErrCodeVal As Long) As String
    ErrStr = UTF8PtrToString(sqlite3_errstr(ErrCodeVal))
End Function

'''' =============================================================
'''' ========================= Statement =========================
'''' =============================================================

#If VBA7 Then
Public Function StmtPrepare16V2(ByVal dbHandle As LongPtr, ByVal SQLQuery As String, ByRef stmtHandle As LongPtr) As Long
#Else
Public Function StmtPrepare16V2(ByVal dbHandle As Long, ByVal SQLQuery As String, ByRef stmtHandle As Long) As Long
#End If
    ' Only compile the first statement in zSql, so *pzTail is left pointing to what remains uncompiled.
    StmtPrepare16V2 = sqlite3_prepare16_v2(dbHandle, StrPtr(SQLQuery), Len(SQLQuery) * 2, stmtHandle, 0)
End Function

'''' =============================================================
'''' ========= Statement column access (0-based indices) =========
'''' =============================================================

#If VBA7 Then
Public Function ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function ColumnName(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    ColumnName = UTF8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
End Function

#If VBA7 Then
Public Function SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim BlobPtr As LongPtr
#Else
Public Function ColumnBlob(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim BlobPtr As Long
#End If
    Dim BlobSize As Long
    BlobPtr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
    Dim BlobBytes() As Byte
    BlobSize = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim BlobBytes(BlobSize - 1)
    RtlMoveMemory VarPtr(BlobBytes(0)), BlobPtr, BlobSize
    ColumnBlob = BlobBytes
End Function

'''' =============================================================
'''' ===================== Statement bindings ====================
'''' =============================================================

#If VBA7 Then
Public Function BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
#Else
Public Function BindText(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
#End If
    BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(Value), -1, SQLITE_TRANSIENT)
End Function

