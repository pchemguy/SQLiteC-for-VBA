VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "SQLiteDBdev.Drafts.SQLite3 Direct"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TSQLite3Constants
    ResultCodes As Scripting.Dictionary '' Result Codes
    ResultCodesEx As Scripting.Dictionary '' Extended Result Codes
    ActionCodes As Scripting.Dictionary '' Authorizer Action Codes
    TextEncodings As Scripting.Dictionary '' Text Encodings
    Datatypes As Scripting.Dictionary '' Fundamental Datatypes
    Config As Scripting.Dictionary '' Configuration Options
    ConnConfig As Scripting.Dictionary '' Database Connection Configuration Options
    FileControl As Scripting.Dictionary '' Standard File Control Opcodes
    FileLocking As Scripting.Dictionary '' File Locking Levels
    MutexTypes As Scripting.Dictionary '' Mutex Types
    FileOpenFlags As Scripting.Dictionary '' Flags For File Open Operations
    PrepareFlags As Scripting.Dictionary '' Prepare Flags
    StmtScanStatus As Scripting.Dictionary '' Prepared Statement Scan Status Opcodes
    TxnState As Scripting.Dictionary '' Allowed return values from [sqlite3_txn_state()]
    RunTimeLimits As Scripting.Dictionary '' Run-Time Limit Categories
    ConnStatus As Scripting.Dictionary '' Status Parameters for database connections
    StmtStatus As Scripting.Dictionary '' Status Parameters for prepared statements
End Type
Private this As TSQLite3Constants


'@Ignore ProcedureNotUsed
Public Function Create() As SQLite3Constants
    Dim Instance As SQLite3Constants
    Set Instance = New SQLite3Constants
    Instance.Init
    Set Create = Instance
End Function


Public Sub Init()
    ResultCodes
    ResultCodesEx
    ActionCodes
    TextEncodings
    Datatypes
    Config
    ConnConfig
    FileControl
    FileLocking
    MutexTypes
    FileOpenFlags
    PrepareFlags
    StmtScanStatus
    TxnState
    RunTimeLimits
    ConnStatus
    StmtStatus
End Sub


Private Sub ResultCodes()
    Set this.ResultCodes = New Scripting.Dictionary
    this.ResultCodes.CompareMode = TextCompare
    With this.ResultCodes
        .Add "SQLITE_OK", 0             '' /* Successful result */
                                        '' /* ===== beginning-of-error-codes ===== */
        .Add "SQLITE_ERROR", 1          '' /* Generic error */
        .Add "SQLITE_INTERNAL", 2       '' /* Internal logic error in SQLite */
        .Add "SQLITE_PERM", 3           '' /* Access permission denied */
        .Add "SQLITE_ABORT", 4          '' /* Callback routine requested an abort */
        .Add "SQLITE_BUSY", 5           '' /* The database file is locked */
        .Add "SQLITE_LOCKED", 6         '' /* A table in the database is locked */
        .Add "SQLITE_NOMEM", 7          '' /* A malloc() failed */
        .Add "SQLITE_READONLY", 8       '' /* Attempt to write a readonly database */
        .Add "SQLITE_INTERRUPT", 9      '' /* Operation terminated by sqlite3_interrupt()*/
        .Add "SQLITE_IOERR", 10         '' /* Some kind of disk I/O error occurred */
        .Add "SQLITE_CORRUPT", 11       '' /* The database disk image is malformed */
        .Add "SQLITE_NOTFOUND", 12      '' /* Unknown opcode in sqlite3_file_control() */
        .Add "SQLITE_FULL", 13          '' /* Insertion failed because database is full */
        .Add "SQLITE_CANTOPEN", 14      '' /* Unable to open the database file */
        .Add "SQLITE_PROTOCOL", 15      '' /* Database lock protocol error */
        .Add "SQLITE_EMPTY", 16         '' /* Internal use only */
        .Add "SQLITE_SCHEMA", 17        '' /* The database schema changed */
        .Add "SQLITE_TOOBIG", 18        '' /* String or BLOB exceeds size limit */
        .Add "SQLITE_CONSTRAINT", 19    '' /* Abort due to constraint violation */
        .Add "SQLITE_MISMATCH", 20      '' /* Data type mismatch */
        .Add "SQLITE_MISUSE", 21        '' /* Library used incorrectly */
        .Add "SQLITE_NOLFS", 22         '' /* Uses OS features not supported on host */
        .Add "SQLITE_AUTH", 23          '' /* Authorization denied */
        .Add "SQLITE_FORMAT", 24        '' /* Not used */
        .Add "SQLITE_RANGE", 25         '' /* 2nd parameter to sqlite3_bind out of range */
        .Add "SQLITE_NOTADB", 26        '' /* File opened that is not a database file */
        .Add "SQLITE_NOTICE", 27        '' /* Notifications from sqlite3_log() */
        .Add "SQLITE_WARNING", 28       '' /* Warnings from sqlite3_log() */
        .Add "SQLITE_ROW", 100          '' /* sqlite3_step() has another row ready */
        .Add "SQLITE_DONE", 101         '' /* sqlite3_step() has finished executing */
                                        '' /* ===== end-of-error-codes ===== */
    End With
End Sub


Private Sub ResultCodesEx()
    Dim SQLITE_ERROR As Long: SQLITE_ERROR = this.ResultCodes("SQLITE_ERROR")
    Dim SQLITE_IOERR As Long: SQLITE_IOERR = this.ResultCodes("SQLITE_IOERR")
    Dim SQLITE_LOCKED As Long: SQLITE_LOCKED = this.ResultCodes("SQLITE_LOCKED")
    Dim SQLITE_BUSY As Long: SQLITE_BUSY = this.ResultCodes("SQLITE_BUSY")
    Dim SQLITE_CANTOPEN As Long: SQLITE_CANTOPEN = this.ResultCodes("SQLITE_CANTOPEN")
    Dim SQLITE_CORRUPT As Long: SQLITE_CORRUPT = this.ResultCodes("SQLITE_CORRUPT")
    Dim SQLITE_READONLY As Long: SQLITE_READONLY = this.ResultCodes("SQLITE_READONLY")
    Dim SQLITE_ABORT As Long: SQLITE_ABORT = this.ResultCodes("SQLITE_ABORT")
    Dim SQLITE_CONSTRAINT As Long: SQLITE_CONSTRAINT = this.ResultCodes("SQLITE_CONSTRAINT")
    Dim SQLITE_NOTICE As Long: SQLITE_NOTICE = this.ResultCodes("SQLITE_NOTICE")
    Dim SQLITE_WARNING As Long: SQLITE_WARNING = this.ResultCodes("SQLITE_WARNING")
    Dim SQLITE_AUTH As Long: SQLITE_AUTH = this.ResultCodes("SQLITE_AUTH")
    Dim SQLITE_OK As Long: SQLITE_OK = this.ResultCodes("SQLITE_OK")
    
    Set this.ResultCodesEx = New Scripting.Dictionary
    this.ResultCodesEx.CompareMode = TextCompare
    With this.ResultCodesEx
        .Add "SQLITE_ERROR_MISSING_COLLSEQ", SQLITE_ERROR + 1 * 2 ^ 8
        .Add "SQLITE_ERROR_RETRY", SQLITE_ERROR + 2 * 2 ^ 8
        .Add "SQLITE_ERROR_SNAPSHOT", SQLITE_ERROR + 3 * 2 ^ 8
        .Add "SQLITE_IOERR_READ", SQLITE_IOERR + 1 * 2 ^ 8
        .Add "SQLITE_IOERR_SHORT_READ", SQLITE_IOERR + 2 * 2 ^ 8
        .Add "SQLITE_IOERR_WRITE", SQLITE_IOERR + 3 * 2 ^ 8
        .Add "SQLITE_IOERR_FSYNC", SQLITE_IOERR + 4 * 2 ^ 8
        .Add "SQLITE_IOERR_DIR_FSYNC", SQLITE_IOERR + 5 * 2 ^ 8
        .Add "SQLITE_IOERR_TRUNCATE", SQLITE_IOERR + 6 * 2 ^ 8
        .Add "SQLITE_IOERR_FSTAT", SQLITE_IOERR + 7 * 2 ^ 8
        .Add "SQLITE_IOERR_UNLOCK", SQLITE_IOERR + 8 * 2 ^ 8
        .Add "SQLITE_IOERR_RDLOCK", SQLITE_IOERR + 9 * 2 ^ 8
        .Add "SQLITE_IOERR_DELETE", SQLITE_IOERR + 10 * 2 ^ 8
        .Add "SQLITE_IOERR_BLOCKED", SQLITE_IOERR + 11 * 2 ^ 8
        .Add "SQLITE_IOERR_NOMEM", SQLITE_IOERR + 12 * 2 ^ 8
        .Add "SQLITE_IOERR_ACCESS", SQLITE_IOERR + 13 * 2 ^ 8
        .Add "SQLITE_IOERR_CHECKRESERVEDLOCK", SQLITE_IOERR + 14 * 2 ^ 8
        .Add "SQLITE_IOERR_LOCK", SQLITE_IOERR + 15 * 2 ^ 8
        .Add "SQLITE_IOERR_CLOSE", SQLITE_IOERR + 16 * 2 ^ 8
        .Add "SQLITE_IOERR_DIR_CLOSE", SQLITE_IOERR + 17 * 2 ^ 8
        .Add "SQLITE_IOERR_SHMOPEN", SQLITE_IOERR + 18 * 2 ^ 8
        .Add "SQLITE_IOERR_SHMSIZE", SQLITE_IOERR + 19 * 2 ^ 8
        .Add "SQLITE_IOERR_SHMLOCK", SQLITE_IOERR + 20 * 2 ^ 8
        .Add "SQLITE_IOERR_SHMMAP", SQLITE_IOERR + 21 * 2 ^ 8
        .Add "SQLITE_IOERR_SEEK", SQLITE_IOERR + 22 * 2 ^ 8
        .Add "SQLITE_IOERR_DELETE_NOENT", SQLITE_IOERR + 23 * 2 ^ 8
        .Add "SQLITE_IOERR_MMAP", SQLITE_IOERR + 24 * 2 ^ 8
        .Add "SQLITE_IOERR_GETTEMPPATH", SQLITE_IOERR + 25 * 2 ^ 8
        .Add "SQLITE_IOERR_CONVPATH", SQLITE_IOERR + 26 * 2 ^ 8
        .Add "SQLITE_IOERR_VNODE", SQLITE_IOERR + 27 * 2 ^ 8
        .Add "SQLITE_IOERR_AUTH", SQLITE_IOERR + 28 * 2 ^ 8
        .Add "SQLITE_IOERR_BEGIN_ATOMIC", SQLITE_IOERR + 29 * 2 ^ 8
        .Add "SQLITE_IOERR_COMMIT_ATOMIC", SQLITE_IOERR + 30 * 2 ^ 8
        .Add "SQLITE_IOERR_ROLLBACK_ATOMIC", SQLITE_IOERR + 31 * 2 ^ 8
        .Add "SQLITE_IOERR_DATA", SQLITE_IOERR + 32 * 2 ^ 8
        .Add "SQLITE_IOERR_CORRUPTFS", SQLITE_IOERR + 33 * 2 ^ 8
        .Add "SQLITE_LOCKED_SHAREDCACHE", SQLITE_LOCKED + 1 * 2 ^ 8
        .Add "SQLITE_LOCKED_VTAB", SQLITE_LOCKED + 2 * 2 ^ 8
        .Add "SQLITE_BUSY_RECOVERY", SQLITE_BUSY + 1 * 2 ^ 8
        .Add "SQLITE_BUSY_SNAPSHOT", SQLITE_BUSY + 2 * 2 ^ 8
        .Add "SQLITE_BUSY_TIMEOUT", SQLITE_BUSY + 3 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_NOTEMPDIR", SQLITE_CANTOPEN + 1 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_ISDIR", SQLITE_CANTOPEN + 2 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_FULLPATH", SQLITE_CANTOPEN + 3 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_CONVPATH", SQLITE_CANTOPEN + 4 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_DIRTYWAL", SQLITE_CANTOPEN + 5 * 2 ^ 8
        .Add "SQLITE_CANTOPEN_SYMLINK", SQLITE_CANTOPEN + 6 * 2 ^ 8
        .Add "SQLITE_CORRUPT_VTAB", SQLITE_CORRUPT + 1 * 2 ^ 8
        .Add "SQLITE_CORRUPT_SEQUENCE", SQLITE_CORRUPT + 2 * 2 ^ 8
        .Add "SQLITE_CORRUPT_INDEX", SQLITE_CORRUPT + 3 * 2 ^ 8
        .Add "SQLITE_READONLY_RECOVERY", SQLITE_READONLY + 1 * 2 ^ 8
        .Add "SQLITE_READONLY_CANTLOCK", SQLITE_READONLY + 2 * 2 ^ 8
        .Add "SQLITE_READONLY_ROLLBACK", SQLITE_READONLY + 3 * 2 ^ 8
        .Add "SQLITE_READONLY_DBMOVED", SQLITE_READONLY + 4 * 2 ^ 8
        .Add "SQLITE_READONLY_CANTINIT", SQLITE_READONLY + 5 * 2 ^ 8
        .Add "SQLITE_READONLY_DIRECTORY", SQLITE_READONLY + 6 * 2 ^ 8
        .Add "SQLITE_ABORT_ROLLBACK", SQLITE_ABORT + 2 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_CHECK", SQLITE_CONSTRAINT + 1 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_COMMITHOOK", SQLITE_CONSTRAINT + 2 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_FOREIGNKEY", SQLITE_CONSTRAINT + 3 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_FUNCTION", SQLITE_CONSTRAINT + 4 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_NOTNULL", SQLITE_CONSTRAINT + 5 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_PRIMARYKEY", SQLITE_CONSTRAINT + 6 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_TRIGGER", SQLITE_CONSTRAINT + 7 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_UNIQUE", SQLITE_CONSTRAINT + 8 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_VTAB", SQLITE_CONSTRAINT + 9 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_ROWID", SQLITE_CONSTRAINT + 10 * 2 ^ 8
        .Add "SQLITE_CONSTRAINT_PINNED", SQLITE_CONSTRAINT + 11 * 2 ^ 8
        .Add "SQLITE_NOTICE_RECOVER_WAL", SQLITE_NOTICE + 1 * 2 ^ 8
        .Add "SQLITE_NOTICE_RECOVER_ROLLBACK", SQLITE_NOTICE + 2 * 2 ^ 8
        .Add "SQLITE_WARNING_AUTOINDEX", SQLITE_WARNING + 1 * 2 ^ 8
        .Add "SQLITE_AUTH_USER", SQLITE_AUTH + 1 * 2 ^ 8
        .Add "SQLITE_OK_LOAD_PERMANENTLY", SQLITE_OK + 1 * 2 ^ 8
        .Add "SQLITE_OK_SYMLINK", SQLITE_OK + 2 * 2 ^ 8
    End With
End Sub


Private Sub ActionCodes()
    Set this.ActionCodes = New Scripting.Dictionary
    this.ActionCodes.CompareMode = TextCompare
    With this.ActionCodes
                                                '' /**** 3rd ************ 4th *********/
        .Add "SQLITE_CREATE_INDEX", 1           '' /* Index Name      Table Name      */
        .Add "SQLITE_CREATE_TABLE", 2           '' /* Table Name      NULL            */
        .Add "SQLITE_CREATE_TEMP_INDEX", 3      '' /* Index Name      Table Name      */
        .Add "SQLITE_CREATE_TEMP_TABLE", 4      '' /* Table Name      NULL            */
        .Add "SQLITE_CREATE_TEMP_TRIGGER", 5    '' /* Trigger Name    Table Name      */
        .Add "SQLITE_CREATE_TEMP_VIEW", 6       '' /* View Name       NULL            */
        .Add "SQLITE_CREATE_TRIGGER", 7         '' /* Trigger Name    Table Name      */
        .Add "SQLITE_CREATE_VIEW", 8            '' /* View Name       NULL            */
        .Add "SQLITE_DELETE", 9                 '' /* Table Name      NULL            */
        .Add "SQLITE_DROP_INDEX", 10            '' /* Index Name      Table Name      */
        .Add "SQLITE_DROP_TABLE", 11            '' /* Table Name      NULL            */
        .Add "SQLITE_DROP_TEMP_INDEX", 12       '' /* Index Name      Table Name      */
        .Add "SQLITE_DROP_TEMP_TABLE", 13       '' /* Table Name      NULL            */
        .Add "SQLITE_DROP_TEMP_TRIGGER", 14     '' /* Trigger Name    Table Name      */
        .Add "SQLITE_DROP_TEMP_VIEW", 15        '' /* View Name       NULL            */
        .Add "SQLITE_DROP_TRIGGER", 16          '' /* Trigger Name    Table Name      */
        .Add "SQLITE_DROP_VIEW", 17             '' /* View Name       NULL            */
        .Add "SQLITE_INSERT", 18                '' /* Table Name      NULL            */
        .Add "SQLITE_PRAGMA", 19                '' /* Pragma Name     1st arg or NULL */
        .Add "SQLITE_READ", 20                  '' /* Table Name      Column Name     */
        .Add "SQLITE_SELECT", 21                '' /* NULL            NULL            */
        .Add "SQLITE_TRANSACTION", 22           '' /* Operation       NULL            */
        .Add "SQLITE_UPDATE", 23                '' /* Table Name      Column Name     */
        .Add "SQLITE_ATTACH", 24                '' /* Filename        NULL            */
        .Add "SQLITE_DETACH", 25                '' /* Database Name   NULL            */
        .Add "SQLITE_ALTER_TABLE", 26           '' /* Database Name   Table Name      */
        .Add "SQLITE_REINDEX", 27               '' /* Index Name      NULL            */
        .Add "SQLITE_ANALYZE", 28               '' /* Table Name      NULL            */
        .Add "SQLITE_CREATE_VTABLE", 29         '' /* Table Name      Module Name     */
        .Add "SQLITE_DROP_VTABLE", 30           '' /* Table Name      Module Name     */
        .Add "SQLITE_FUNCTION", 31              '' /* NULL            Function Name   */
        .Add "SQLITE_SAVEPOINT", 32             '' /* Operation       Savepoint Name  */
        .Add "SQLITE_COPY", 0                   '' /* No longer used                  */
        .Add "SQLITE_RECURSIVE", 33             '' /* NULL            NULL            */
    End With
End Sub


Private Sub TextEncodings()
    Set this.TextEncodings = New Scripting.Dictionary
    this.TextEncodings.CompareMode = TextCompare
    With this.TextEncodings
        .Add "SQLITE_UTF8", 1           '' /* IMP: R-37514-35566 */
        .Add "SQLITE_UTF16LE", 2        '' /* IMP: R-03371-37637 */
        .Add "SQLITE_UTF16BE", 3        '' /* IMP: R-51971-34154 */
        .Add "SQLITE_UTF16", 4          '' /* Use native byte order */
        .Add "SQLITE_ANY", 5            '' /* Deprecated */
        .Add "SQLITE_UTF16_ALIGNED", 8  '' /* sqlite3_create_collation only */
    End With
End Sub


Private Sub Datatypes()
    Set this.Datatypes = New Scripting.Dictionary
    this.Datatypes.CompareMode = TextCompare
    With this.Datatypes
        .Add "SQLITE_INTEGER", 1
        .Add "SQLITE_FLOAT", 2
        .Add "SQLITE_BLOB", 4
        .Add "SQLITE_TEXT", 3
        .Add "SQLITE_NULL", 5
    End With
End Sub


Private Sub Config()
    Set this.Config = New Scripting.Dictionary
    this.Config.CompareMode = TextCompare
    With this.Config
        .Add "SQLITE_CONFIG_SINGLETHREAD", 1           '' /* nil */
        .Add "SQLITE_CONFIG_MULTITHREAD", 2            '' /* nil */
        .Add "SQLITE_CONFIG_SERIALIZED", 3             '' /* nil */
        .Add "SQLITE_CONFIG_MALLOC", 4                 '' /* sqlite3_mem_methods* */
        .Add "SQLITE_CONFIG_GETMALLOC", 5              '' /* sqlite3_mem_methods* */
        .Add "SQLITE_CONFIG_SCRATCH", 6                '' /* No longer used */
        .Add "SQLITE_CONFIG_PAGECACHE", 7              '' /* void*, int sz, int N */
        .Add "SQLITE_CONFIG_HEAP", 8                   '' /* void*, int nByte, int min */
        .Add "SQLITE_CONFIG_MEMSTATUS", 9              '' /* boolean */
        .Add "SQLITE_CONFIG_MUTEX", 10                 '' /* sqlite3_mutex_methods* */
        .Add "SQLITE_CONFIG_GETMUTEX", 11              '' /* sqlite3_mutex_methods* */
        .Add "SQLITE_CONFIG_CHUNKALLOC", 12            '' /* which is now unused. */
        .Add "SQLITE_CONFIG_LOOKASIDE", 13             '' /* int int */
        .Add "SQLITE_CONFIG_PCACHE", 14                '' /* no-op */
        .Add "SQLITE_CONFIG_GETPCACHE", 15             '' /* no-op */
        .Add "SQLITE_CONFIG_LOG", 16                   '' /* xFunc, void* */
        .Add "SQLITE_CONFIG_URI", 17                   '' /* int */
        .Add "SQLITE_CONFIG_PCACHE2", 18               '' /* sqlite3_pcache_methods2* */
        .Add "SQLITE_CONFIG_GETPCACHE2", 19            '' /* sqlite3_pcache_methods2* */
        .Add "SQLITE_CONFIG_COVERING_INDEX_SCAN", 20   '' /* int */
        .Add "SQLITE_CONFIG_SQLLOG", 21                '' /* xSqllog, void* */
        .Add "SQLITE_CONFIG_MMAP_SIZE", 22             '' /* sqlite3_int64, sqlite3_int64 */
        .Add "SQLITE_CONFIG_WIN32_HEAPSIZE", 23        '' /* int nByte */
        .Add "SQLITE_CONFIG_PCACHE_HDRSZ", 24          '' /* int *psz */
        .Add "SQLITE_CONFIG_PMASZ", 25                 '' /* unsigned int szPma */
        .Add "SQLITE_CONFIG_STMTJRNL_SPILL", 26        '' /* int nByte */
        .Add "SQLITE_CONFIG_SMALL_MALLOC", 27          '' /* boolean */
        .Add "SQLITE_CONFIG_SORTERREF_SIZE", 28        '' /* int nByte */
        .Add "SQLITE_CONFIG_MEMDB_MAXSIZE", 29         '' /* sqlite3_int64 */
    End With
End Sub


Private Sub ConnConfig()
    Set this.ConnConfig = New Scripting.Dictionary
    this.ConnConfig.CompareMode = TextCompare
    With this.ConnConfig
        .Add "SQLITE_DBCONFIG_MAINDBNAME", 1000             '' /* const char* */
        .Add "SQLITE_DBCONFIG_LOOKASIDE", 1001              '' /* void* int int */
        .Add "SQLITE_DBCONFIG_ENABLE_FKEY", 1002            '' /* int int* */
        .Add "SQLITE_DBCONFIG_ENABLE_TRIGGER", 1003         '' /* int int* */
        .Add "SQLITE_DBCONFIG_ENABLE_FTS3_TOKENIZER", 1004  '' /* int int* */
        .Add "SQLITE_DBCONFIG_ENABLE_LOAD_EXTENSION", 1005  '' /* int int* */
        .Add "SQLITE_DBCONFIG_NO_CKPT_ON_CLOSE", 1006       '' /* int int* */
        .Add "SQLITE_DBCONFIG_ENABLE_QPSG", 1007            '' /* int int* */
        .Add "SQLITE_DBCONFIG_TRIGGER_EQP", 1008            '' /* int int* */
        .Add "SQLITE_DBCONFIG_RESET_DATABASE", 1009         '' /* int int* */
        .Add "SQLITE_DBCONFIG_DEFENSIVE", 1010              '' /* int int* */
        .Add "SQLITE_DBCONFIG_WRITABLE_SCHEMA", 1011        '' /* int int* */
        .Add "SQLITE_DBCONFIG_LEGACY_ALTER_TABLE", 1012     '' /* int int* */
        .Add "SQLITE_DBCONFIG_DQS_DML", 1013                '' /* int int* */
        .Add "SQLITE_DBCONFIG_DQS_DDL", 1014                '' /* int int* */
        .Add "SQLITE_DBCONFIG_ENABLE_VIEW", 1015            '' /* int int* */
        .Add "SQLITE_DBCONFIG_LEGACY_FILE_FORMAT", 1016     '' /* int int* */
        .Add "SQLITE_DBCONFIG_TRUSTED_SCHEMA", 1017         '' /* int int* */
        .Add "SQLITE_DBCONFIG_MAX", 1017                    '' /* Largest DBCONFIG */
    End With
End Sub


Private Sub FileControl()
    Set this.FileControl = New Scripting.Dictionary
    this.FileControl.CompareMode = TextCompare
    With this.FileControl
        .Add "SQLITE_FCNTL_LOCKSTATE", 1
        .Add "SQLITE_FCNTL_GET_LOCKPROXYFILE", 2
        .Add "SQLITE_FCNTL_SET_LOCKPROXYFILE", 3
        .Add "SQLITE_FCNTL_LAST_ERRNO", 4
        .Add "SQLITE_FCNTL_SIZE_HINT", 5
        .Add "SQLITE_FCNTL_CHUNK_SIZE", 6
        .Add "SQLITE_FCNTL_FILE_POINTER", 7
        .Add "SQLITE_FCNTL_SYNC_OMITTED", 8
        .Add "SQLITE_FCNTL_WIN32_AV_RETRY", 9
        .Add "SQLITE_FCNTL_PERSIST_WAL", 10
        .Add "SQLITE_FCNTL_OVERWRITE", 11
        .Add "SQLITE_FCNTL_VFSNAME", 12
        .Add "SQLITE_FCNTL_POWERSAFE_OVERWRITE", 13
        .Add "SQLITE_FCNTL_PRAGMA", 14
        .Add "SQLITE_FCNTL_BUSYHANDLER", 15
        .Add "SQLITE_FCNTL_TEMPFILENAME", 16
        .Add "SQLITE_FCNTL_MMAP_SIZE", 18
        .Add "SQLITE_FCNTL_TRACE", 19
        .Add "SQLITE_FCNTL_HAS_MOVED", 20
        .Add "SQLITE_FCNTL_SYNC", 21
        .Add "SQLITE_FCNTL_COMMIT_PHASETWO", 22
        .Add "SQLITE_FCNTL_WIN32_SET_HANDLE", 23
        .Add "SQLITE_FCNTL_WAL_BLOCK", 24
        .Add "SQLITE_FCNTL_ZIPVFS", 25
        .Add "SQLITE_FCNTL_RBU", 26
        .Add "SQLITE_FCNTL_VFS_POINTER", 27
        .Add "SQLITE_FCNTL_JOURNAL_POINTER", 28
        .Add "SQLITE_FCNTL_WIN32_GET_HANDLE", 29
        .Add "SQLITE_FCNTL_PDB", 30
        .Add "SQLITE_FCNTL_BEGIN_ATOMIC_WRITE", 31
        .Add "SQLITE_FCNTL_COMMIT_ATOMIC_WRITE", 32
        .Add "SQLITE_FCNTL_ROLLBACK_ATOMIC_WRITE", 33
        .Add "SQLITE_FCNTL_LOCK_TIMEOUT", 34
        .Add "SQLITE_FCNTL_DATA_VERSION", 35
        .Add "SQLITE_FCNTL_SIZE_LIMIT", 36
        .Add "SQLITE_FCNTL_CKPT_DONE", 37
        .Add "SQLITE_FCNTL_RESERVE_BYTES", 38
        .Add "SQLITE_FCNTL_CKPT_START", 39
    End With
End Sub


Private Sub FileLocking()
    Set this.FileLocking = New Scripting.Dictionary
    this.FileLocking.CompareMode = TextCompare
    With this.FileLocking
        .Add "SQLITE_LOCK_NONE", 0
        .Add "SQLITE_LOCK_SHARED", 1
        .Add "SQLITE_LOCK_RESERVED", 2
        .Add "SQLITE_LOCK_PENDING", 3
        .Add "SQLITE_LOCK_EXCLUSIVE", 4
    End With
End Sub



Private Sub MutexTypes()
    Set this.MutexTypes = New Scripting.Dictionary
    this.MutexTypes.CompareMode = TextCompare
    With this.MutexTypes
        .Add "SQLITE_MUTEX_FAST", 0
        .Add "SQLITE_MUTEX_RECURSIVE", 1
        .Add "SQLITE_MUTEX_STATIC_MAIN", 2
        .Add "SQLITE_MUTEX_STATIC_MEM", 3       '' /* sqlite3_malloc() */
        .Add "SQLITE_MUTEX_STATIC_MEM2", 4      '' /* NOT USED */
        .Add "SQLITE_MUTEX_STATIC_OPEN", 4      '' /* sqlite3BtreeOpen() */
        .Add "SQLITE_MUTEX_STATIC_PRNG", 5      '' /* sqlite3_randomness() */
        .Add "SQLITE_MUTEX_STATIC_LRU", 6       '' /* lru page list */
        .Add "SQLITE_MUTEX_STATIC_LRU2", 7      '' /* NOT USED */
        .Add "SQLITE_MUTEX_STATIC_PMEM", 7      '' /* sqlite3PageMalloc() */
        .Add "SQLITE_MUTEX_STATIC_APP1", 8      '' /* For use by application */
        .Add "SQLITE_MUTEX_STATIC_APP2", 9      '' /* For use by application */
        .Add "SQLITE_MUTEX_STATIC_APP3", 10     '' /* For use by application */
        .Add "SQLITE_MUTEX_STATIC_VFS1", 11     '' /* For use by built-in VFS */
        .Add "SQLITE_MUTEX_STATIC_VFS2", 12     '' /* For use by extension VFS */
        .Add "SQLITE_MUTEX_STATIC_VFS3", 13     '' /* For use by application VFS */
    End With
End Sub


Private Sub FileOpenFlags()
    Set this.FileOpenFlags = New Scripting.Dictionary
    this.FileOpenFlags.CompareMode = TextCompare
    With this.FileOpenFlags
        .Add "SQLITE_OPEN_READONLY", 2 ^ 0          '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_READWRITE", 2 ^ 1         '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_CREATE", 2 ^ 2            '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_DELETEONCLOSE", 2 ^ 3     '' /* VFS only */
        .Add "SQLITE_OPEN_EXCLUSIVE", 2 ^ 4         '' /* VFS only */
        .Add "SQLITE_OPEN_AUTOPROXY", 2 ^ 5         '' /* VFS only */
        .Add "SQLITE_OPEN_URI", 2 ^ 6               '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_MEMORY", 2 ^ 7            '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_MAIN_DB", 2 ^ 8           '' /* VFS only */
        .Add "SQLITE_OPEN_TEMP_DB", 2 ^ 9           '' /* VFS only */
        .Add "SQLITE_OPEN_TRANSIENT_DB", 2 ^ 10     '' /* VFS only */
        .Add "SQLITE_OPEN_MAIN_JOURNAL", 2 ^ 11     '' /* VFS only */
        .Add "SQLITE_OPEN_TEMP_JOURNAL", 2 ^ 12     '' /* VFS only */
        .Add "SQLITE_OPEN_SUBJOURNAL", 2 ^ 13       '' /* VFS only */
        .Add "SQLITE_OPEN_SUPER_JOURNAL", 2 ^ 14    '' /* VFS only */
        .Add "SQLITE_OPEN_NOMUTEX", 2 ^ 15          '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_FULLMUTEX", 2 ^ 16        '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_SHAREDCACHE", 2 ^ 17      '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_PRIVATECACHE", 2 ^ 18     '' /* Ok for sqlite3_open_v2() */
        .Add "SQLITE_OPEN_WAL", 2 ^ 19              '' /* VFS only */
        .Add "SQLITE_OPEN_NOFOLLOW", 2 ^ 24         '' /* Ok for sqlite3_open_v2() */
    End With
End Sub


Private Sub PrepareFlags()
    Set this.PrepareFlags = New Scripting.Dictionary
    this.PrepareFlags.CompareMode = TextCompare
    With this.PrepareFlags
        .Add "SQLITE_PREPARE_PERSISTENT", 1
        .Add "SQLITE_PREPARE_NORMALIZE", 2
        .Add "SQLITE_PREPARE_NO_VTAB", 4
    End With
End Sub

    
Private Sub StmtScanStatus()
    Set this.StmtScanStatus = New Scripting.Dictionary
    this.StmtScanStatus.CompareMode = TextCompare
    With this.StmtScanStatus
        .Add "SQLITE_SCANSTAT_NLOOP", 0
        .Add "SQLITE_SCANSTAT_NVISIT", 1
        .Add "SQLITE_SCANSTAT_EST", 2
        .Add "SQLITE_SCANSTAT_NAME", 3
        .Add "SQLITE_SCANSTAT_EXPLAIN", 4
        .Add "SQLITE_SCANSTAT_SELECTID", 5
    End With
End Sub


Private Sub TxnState()
    Set this.TxnState = New Scripting.Dictionary
    this.TxnState.CompareMode = TextCompare
    With this.TxnState
        .Add "SQLITE_TXN_NONE", 0
        .Add "SQLITE_TXN_READ", 1
        .Add "SQLITE_TXN_WRITE", 2
    End With
End Sub


Private Sub RunTimeLimits()
    Set this.RunTimeLimits = New Scripting.Dictionary
    this.RunTimeLimits.CompareMode = TextCompare
    With this.RunTimeLimits
        .Add "SQLITE_LIMIT_LENGTH", 0
        .Add "SQLITE_LIMIT_SQL_LENGTH", 1
        .Add "SQLITE_LIMIT_COLUMN", 2
        .Add "SQLITE_LIMIT_EXPR_DEPTH", 3
        .Add "SQLITE_LIMIT_COMPOUND_SELECT", 4
        .Add "SQLITE_LIMIT_VDBE_OP", 5
        .Add "SQLITE_LIMIT_FUNCTION_ARG", 6
        .Add "SQLITE_LIMIT_ATTACHED", 7
        .Add "SQLITE_LIMIT_LIKE_PATTERN_LENGTH", 8
        .Add "SQLITE_LIMIT_VARIABLE_NUMBER", 9
        .Add "SQLITE_LIMIT_TRIGGER_DEPTH", 10
        .Add "SQLITE_LIMIT_WORKER_THREADS", 11
    End With
End Sub


Private Sub ConnStatus()
    Set this.ConnStatus = New Scripting.Dictionary
    this.ConnStatus.CompareMode = TextCompare
    With this.ConnStatus
        .Add "SQLITE_DBSTATUS_LOOKASIDE_USED", 0
        .Add "SQLITE_DBSTATUS_CACHE_USED", 1
        .Add "SQLITE_DBSTATUS_SCHEMA_USED", 2
        .Add "SQLITE_DBSTATUS_STMT_USED", 3
        .Add "SQLITE_DBSTATUS_LOOKASIDE_HIT", 4
        .Add "SQLITE_DBSTATUS_LOOKASIDE_MISS_SIZE", 5
        .Add "SQLITE_DBSTATUS_LOOKASIDE_MISS_FULL", 6
        .Add "SQLITE_DBSTATUS_CACHE_HIT", 7
        .Add "SQLITE_DBSTATUS_CACHE_MISS", 8
        .Add "SQLITE_DBSTATUS_CACHE_WRITE", 9
        .Add "SQLITE_DBSTATUS_DEFERRED_FKS", 10
        .Add "SQLITE_DBSTATUS_CACHE_USED_SHARED", 11
        .Add "SQLITE_DBSTATUS_CACHE_SPILL", 12
        .Add "SQLITE_DBSTATUS_MAX", 12                  '' /* Largest defined DBSTATUS */
    End With
End Sub


Private Sub StmtStatus()
    Set this.StmtStatus = New Scripting.Dictionary
    this.StmtStatus.CompareMode = TextCompare
    With this.StmtStatus
        .Add "SQLITE_STMTSTATUS_FULLSCAN_STEP", 1
        .Add "SQLITE_STMTSTATUS_SORT", 2
        .Add "SQLITE_STMTSTATUS_AUTOINDEX", 3
        .Add "SQLITE_STMTSTATUS_VM_STEP", 4
        .Add "SQLITE_STMTSTATUS_REPREPARE", 5
        .Add "SQLITE_STMTSTATUS_RUN", 6
        .Add "SQLITE_STMTSTATUS_MEMUSED", 99
    End With
End Sub
