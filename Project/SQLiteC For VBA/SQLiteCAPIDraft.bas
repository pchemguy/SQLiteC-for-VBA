Attribute VB_Name = "SQLiteCAPIDraft"
'@Folder "SQLiteC For VBA"
'@IgnoreModule EmptyModule
Option Explicit


'
'
'#If VBA7 Then
'
'' SQLite library version
'Public Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
'
'' Database connections
'Public Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
'Public Declare PtrSafe Function DbClose Lib "SQLite3" Alias "sqlite3_close" (ByVal hDb As LongPtr) As Long
'
'' Database connection error info
'Public Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
'Public Declare PtrSafe Function sqlite3_errstr Lib "SQLite3" (ByVal ErrorCode As Long) As LongPtr ' PtrUtf8String
'Public Declare PtrSafe Function ErrCode Lib "SQLite3" Alias "sqlite3_errcode" (ByVal hDb As LongPtr) As Long
'Public Declare PtrSafe Function ErrCodeExtended Lib "SQLite3" Alias "sqlite3_extended_errcode" (ByVal hDb As LongPtr) As Long
'
'' Database connection change counts
'Public Declare PtrSafe Function Changes Lib "SQLite3" Alias "sqlite3_changes" (ByVal hDb As LongPtr) As Long
'Public Declare PtrSafe Function ChangesTotal Lib "SQLite3" Alias "sqlite3_total_changes" (ByVal hDb As LongPtr) As Long
'
'' Statements
'Public Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
'    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
'Public Declare PtrSafe Function StmtStep Lib "SQLite3" Alias "sqlite3_step" (ByVal hStmt As LongPtr) As Long
'Public Declare PtrSafe Function StmtReset Lib "SQLite3" Alias "sqlite3_reset" (ByVal hStmt As LongPtr) As Long
'Public Declare PtrSafe Function StmtFinalize Lib "SQLite3" Alias "sqlite3_finalize" (ByVal hStmt As LongPtr) As Long
'
'' Statement column access (0-based indices)
'Public Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
'Public Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
'
'Public Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
'Public Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
'Public Declare PtrSafe Function ColumnDouble Lib "SQLite3" Alias "sqlite3_column_double" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
'Public Declare PtrSafe Function ColumnInt32 Lib "SQLite3" Alias "sqlite3_column_int" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
'Public Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
'Public Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
'Public Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value
'
'' Statement parameter binding (1-based indices!)
'Public Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
'Public Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
'Public Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
'Public Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Double) As Long
'Public Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Long) As Long
'Public Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As LongLong) As Long
'Public Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
'
''Backup
'Public Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
'Public Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
'Public Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'Public Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'
'#Else
'
'' SQLite library version
'Public Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
'
'' Database connections
'Public Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
'Public Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
'Public Declare Function DbClose Lib "SQLite3" Alias "sqlite3_close" (ByVal hDb As Long) As Long
'
'' Database connection error info
'Public Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
'Public Declare Function sqlite3_errstr Lib "SQLite3" (ByVal ErrorCode As Long) As Long ' PtrUtf8String
'Public Declare Function ErrCode Lib "SQLite3" Alias "sqlite3_errcode" (ByVal hDb As Long) As Long
'Public Declare Function ErrCodeExtended Lib "SQLite3" Alias "sqlite3_extended_errcode" (ByVal hDb As Long) As Long
'
'' Database connection change counts
'Public Declare Function Changes Lib "SQLite3" Alias "sqlite3_changes" (ByVal hDb As Long) As Long
'Public Declare Function ChangesTotal Lib "SQLite3" Alias "sqlite3_total_changes" (ByVal hDb As Long) As Long
'
'' Statements
'Public Declare Function sqlite3_prepare16_v2 Lib "SQLite3" _
'    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
'Public Declare Function StmtStep Lib "SQLite3" Alias "sqlite3_step" (ByVal hStmt As Long) As Long
'Public Declare Function StmtReset Lib "SQLite3" Alias "sqlite3_reset" (ByVal hStmt As Long) As Long
'Public Declare Function StmtFinalize Lib "SQLite3" Alias "sqlite3_finalize" (ByVal hStmt As Long) As Long
'
'' Statement column access (0-based indices)
'Public Declare Function ColumnCount Lib "SQLite3" Alias "sqlite3_column_count" (ByVal hStmt As Long) As Long
'Public Declare Function ColumnType Lib "SQLite3" Alias "sqlite3_column_type" (ByVal hStmt As Long, ByVal iCol As Long) As Long
'Public Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
'
'Public Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
'Public Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
'Public Declare Function ColumnDouble Lib "SQLite3" Alias "sqlite3_column_double" (ByVal hStmt As Long, ByVal iCol As Long) As Double
'Public Declare Function ColumnInt32 Lib "SQLite3" Alias "sqlite3_column_int" (ByVal hStmt As Long, ByVal iCol As Long) As Long
'Public Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency
'Public Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
'Public Declare Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value
'
'' Statement parameter binding (1-based indices!)
'Public Declare Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As Long) As Long
'Public Declare Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
'Public Declare Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As Long, ByVal paramName As Long) As Long
'Public Declare Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
'Public Declare Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
'Public Declare Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
'Public Declare Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Double) As Long
'Public Declare Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Long) As Long
'Public Declare Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Currency) As Long
'Public Declare Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
'Public Declare Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
'Public Declare Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
'Public Declare Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As Long) As Long
'
''Backup
'Public Declare Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
'Public Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
'Public Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
'Public Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
'Public Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
'Public Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
'
'#End If
'
'
''''' =============================================================
''''' =========================== Meta ============================
''''' =============================================================
'
'Public Function LibVersion() As String
'    LibVersion = Utf8PtrToString(sqlite3_libversion())
'End Function
'
''''' =============================================================
''''' ==================== Database connection ====================
''''' =============================================================
'
'#If VBA7 Then
'Public Function DbOpen16(ByVal FileName As String, ByRef dbHandle As LongPtr) As Long
'#Else
'Public Function DbOpen16(ByVal FileName As String, ByRef dbHandle As Long) As Long
'#End If
'    DbOpen16 = sqlite3_open16(StrPtr(FileName), dbHandle)
'End Function
'
'#If VBA7 Then
'Public Function DbOpenV2(ByVal FileName As String, ByRef dbHandle As LongPtr, ByVal Flags As Long, ByVal vfsName As String) As Long
'#Else
'Public Function DbOpenV2(ByVal FileName As String, ByRef dbHandle As Long, ByVal Flags As Long, ByVal vfsName As String) As Long
'#End If
'    Dim FileNameBytes() As Byte
'    Dim vfsNameBytes() As Byte
'    FileNameBytes = StringToUtf8Bytes(FileName)
'    If Len(vfsName) = 0 Then
'        DbOpenV2 = sqlite3_open_v2(VarPtr(FileNameBytes(0)), dbHandle, Flags, 0)
'    Else
'        vfsNameBytes = StringToUtf8Bytes(vfsName)
'        DbOpenV2 = sqlite3_open_v2(VarPtr(FileNameBytes(0)), dbHandle, Flags, VarPtr(vfsNameBytes(0)))
'    End If
'End Function
'
''''' =============================================================
''''' ===================== Error Information =====================
''''' =============================================================
'
'#If VBA7 Then
'Public Function ErrMsg(ByVal dbHandle As LongPtr) As String
'#Else
'Public Function ErrMsg(ByVal dbHandle As Long) As String
'#End If
'    ErrMsg = Utf8PtrToString(sqlite3_errmsg(dbHandle))
'End Function
'
'Public Function ErrStr(ByVal ErrCodeVal As Long) As String
'    ErrStr = Utf8PtrToString(sqlite3_errstr(ErrCodeVal))
'End Function
'
''''' =============================================================
''''' ========================= Statement =========================
''''' =============================================================
'
'#If VBA7 Then
'Public Function StmtPrepare16V2(ByVal dbHandle As LongPtr, ByVal SQLQuery As String, ByRef stmtHandle As LongPtr) As Long
'#Else
'Public Function StmtPrepare16V2(ByVal dbHandle As Long, ByVal SQLQuery As String, ByRef stmtHandle As Long) As Long
'#End If
'    ' Only compile the first statement in zSql, so *pzTail is left pointing to what remains uncompiled.
'    StmtPrepare16V2 = sqlite3_prepare16_v2(dbHandle, StrPtr(SQLQuery), Len(SQLQuery) * 2, stmtHandle, 0)
'End Function
'
''''' =============================================================
''''' ========= Statement column access (0-based indices) =========
''''' =============================================================
'
'#If VBA7 Then
'Public Function ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
'#Else
'Public Function ColumnName(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
'#End If
'    ColumnName = Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
'End Function
'
'#If VBA7 Then
'Public Function SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
'    Dim BlobPtr As LongPtr
'#Else
'Public Function ColumnBlob(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Byte()
'    Dim BlobPtr As Long
'#End If
'    Dim BlobSize As Long
'    BlobPtr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
'    Dim BlobBytes() As Byte
'    BlobSize = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
'    ReDim BlobBytes(BlobSize - 1)
'    RtlMoveMemory VarPtr(BlobBytes(0)), BlobPtr, BlobSize
'    ColumnBlob = BlobBytes
'End Function
'
''''' =============================================================
''''' ===================== Statement bindings ====================
''''' =============================================================
'
'#If VBA7 Then
'Public Function BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
'#Else
'Public Function BindText(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
'#End If
'    BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(Value), -1, SQLITE_TRANSIENT)
'End Function
'
'
'


