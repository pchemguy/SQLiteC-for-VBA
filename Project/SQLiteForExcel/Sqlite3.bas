Attribute VB_Name = "Sqlite3"
'@Folder "SQLiteForExcel"
'@IgnoreModule
Option Explicit

'
' Notes:
' Microsoft uses UTF-16, little endian byte order.

#If WIN64 Then
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
#Else
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
#End If
'=====================================================================================
' SQLite StdCall Imports
'-----------------------
#If WIN64 Then
' SQLite library version
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
' Database connections
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf16String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long

' Statements
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

' Statement column access (0-based indices)
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString

Private Declare PtrSafe Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrData
Private Declare PtrSafe Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Double
Private Declare PtrSafe Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongLong
Private Declare PtrSafe Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrWString
Private Declare PtrSafe Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Double) As Long
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As LongLong) As Long
Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long

'Backup
Private Declare PtrSafe Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else

' SQLite library version
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
' Database connections
Private Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3" (ByVal hDb As Long) As Long
' Database connection error info
Private Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf16String
Private Declare Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As Long) As Long

' Statements
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As Long
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As Long) As Long

' Statement column access (0-based indices)
Private Declare Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_type Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_name16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString

Private Declare Function sqlite3_column_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrData
Private Declare Function sqlite3_column_bytes Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_bytes16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_double Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Double
Private Declare Function sqlite3_column_int Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long
Private Declare Function sqlite3_column_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Currency ' UNTESTED ....?
Private Declare Function sqlite3_column_text Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrWString
Private Declare Function sqlite3_column_value Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrSqlite3Value

' Statement parameter binding (1-based indices!)
Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As Long, ByVal paramName As Long) As Long
Private Declare Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Double) As Long
Private Declare Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal Value As Currency) As Long ' UNTESTED ....?
Private Declare Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As Long, ByVal paramIndex As Long, ByVal pSqlite3Value As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As Long) As Long

'Backup
Private Declare Function sqlite3_sleep Lib "SQLite3" (ByVal msToSleep As Long) As Long
Private Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
#End If
'=====================================================================================




'=====================================================================================
' SQLite library version

Public Function SQLite3LibVersion() As String
    SQLite3LibVersion = UTFlib.Utf8PtrToString(sqlite3_libversion())
End Function


'=====================================================================================
' Database connections
#If WIN64 Then
Public Function SQLite3Open(ByVal FileName As String, ByRef dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Open(ByVal FileName As String, ByRef dbHandle As Long) As Long
#End If
    SQLite3Open = sqlite3_open16(StrPtr(FileName), dbHandle)
End Function

#If WIN64 Then
Public Function SQLite3OpenV2(ByVal FileName As String, ByRef dbHandle As LongPtr, ByVal Flags As Long, ByVal vfsName As String) As Long
#Else
Public Function SQLite3OpenV2(ByVal FileName As String, ByRef dbHandle As Long, ByVal Flags As Long, ByVal vfsName As String) As Long
#End If

    Dim bufFileName() As Byte
    Dim bufVfsName() As Byte
    bufFileName = UTFlib.StringToUtf8Bytes(FileName)
    If vfsName = Empty Then
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, Flags, 0)
    Else
        bufVfsName = UTFlib.StringToUtf8Bytes(vfsName)
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, Flags, VarPtr(bufVfsName(0)))
    End If

End Function

#If WIN64 Then
Public Function SQLite3Close(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Close(ByVal dbHandle As Long) As Long
#End If
    SQLite3Close = sqlite3_close(dbHandle)
End Function

'=====================================================================================
' Error information

#If WIN64 Then
Public Function SQLite3ErrMsg(ByVal dbHandle As LongPtr) As String
#Else
Public Function SQLite3ErrMsg(ByVal dbHandle As Long) As String
#End If
    SQLite3ErrMsg = UTFlib.Utf8PtrToString(sqlite3_errmsg(dbHandle))
End Function

#If WIN64 Then
Public Function SQLite3ErrCode(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3ErrCode(ByVal dbHandle As Long) As Long
#End If
    SQLite3ErrCode = sqlite3_errcode(dbHandle)
End Function

#If WIN64 Then
Public Function SQLite3ExtendedErrCode(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3ExtendedErrCode(ByVal dbHandle As Long) As Long
#End If
    SQLite3ExtendedErrCode = sqlite3_extended_errcode(dbHandle)
End Function

'=====================================================================================
' Change Counts

#If WIN64 Then
Public Function SQLite3Changes(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Changes(ByVal dbHandle As Long) As Long
#End If
    SQLite3Changes = sqlite3_changes(dbHandle)
End Function

#If WIN64 Then
Public Function SQLite3TotalChanges(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3TotalChanges(ByVal dbHandle As Long) As Long
#End If
    SQLite3TotalChanges = sqlite3_total_changes(dbHandle)
End Function

'=====================================================================================
' Statements

#If WIN64 Then
Public Function SQLite3PrepareV2(ByVal dbHandle As LongPtr, ByVal sql As String, ByRef stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3PrepareV2(ByVal dbHandle As Long, ByVal sql As String, ByRef stmtHandle As Long) As Long
#End If
    ' Only the first statement (up to ';') is prepared. Currently we don't retrieve the 'tail' pointer.
    SQLite3PrepareV2 = sqlite3_prepare16_v2(dbHandle, StrPtr(sql), Len(sql) * 2, stmtHandle, 0)
End Function

#If WIN64 Then
Public Function SQLite3Step(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Step(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Step = sqlite3_step(stmtHandle)
End Function

#If WIN64 Then
Public Function SQLite3Reset(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Reset(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Reset = sqlite3_reset(stmtHandle)
End Function

#If WIN64 Then
Public Function SQLite3Finalize(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3Finalize(ByVal stmtHandle As Long) As Long
#End If
    SQLite3Finalize = sqlite3_finalize(stmtHandle)
End Function

'=====================================================================================
' Statement column access (0-based indices)

#If WIN64 Then
Public Function SQLite3ColumnCount(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3ColumnCount(ByVal stmtHandle As Long) As Long
#End If
    SQLite3ColumnCount = sqlite3_column_count(stmtHandle)
End Function

#If WIN64 Then
Public Function SQLite3ColumnType(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
#Else
Public Function SQLite3ColumnType(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
#End If
    SQLite3ColumnType = sqlite3_column_type(stmtHandle, ZeroBasedColIndex)
End Function

#If WIN64 Then
Public Function SQLite3ColumnName(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function SQLite3ColumnName(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    SQLite3ColumnName = UTFlib.Utf8PtrToString(sqlite3_column_name(stmtHandle, ZeroBasedColIndex))
End Function

#If WIN64 Then
Public Function SQLite3ColumnDouble(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Double
#Else
Public Function SQLite3ColumnDouble(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Double
#End If
    SQLite3ColumnDouble = sqlite3_column_double(stmtHandle, ZeroBasedColIndex)
End Function

#If WIN64 Then
Public Function SQLite3ColumnInt32(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Long
#Else
Public Function SQLite3ColumnInt32(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Long
#End If
    SQLite3ColumnInt32 = sqlite3_column_int(stmtHandle, ZeroBasedColIndex)
End Function

#If WIN64 Then
Public Function SQLite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function SQLite3ColumnText(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    SQLite3ColumnText = UTFlib.Utf8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
End Function

#If WIN64 Then
Public Function SQLite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date
#Else
Public Function SQLite3ColumnDate(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Date
#End If
    SQLite3ColumnDate = FromJulianDay(sqlite3_column_double(stmtHandle, ZeroBasedColIndex))
End Function

#If WIN64 Then
Public Function SQLite3ColumnBlob(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As LongPtr
#Else
Public Function SQLite3ColumnBlob(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Byte()
    Dim ptr As Long
#End If

    Dim Length As Long
    Dim buf() As Byte
    
    ptr = sqlite3_column_blob(stmtHandle, ZeroBasedColIndex)
    Length = sqlite3_column_bytes(stmtHandle, ZeroBasedColIndex)
    ReDim buf(Length - 1)
    RtlMoveMemory VarPtr(buf(0)), ptr, Length
    SQLite3ColumnBlob = buf
End Function
'=====================================================================================
' Statement bindings

#If WIN64 Then
Public Function SQLite3BindText(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
#Else
Public Function SQLite3BindText(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As String) As Long
#End If
    SQLite3BindText = sqlite3_bind_text16(stmtHandle, OneBasedParamIndex, StrPtr(Value), -1, SQLITE_TRANSIENT)
End Function

#If WIN64 Then
Public Function SQLite3BindDouble(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Double) As Long
#Else
Public Function SQLite3BindDouble(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Double) As Long
#End If
    SQLite3BindDouble = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, Value)
End Function

#If WIN64 Then
Public Function SQLite3BindInt32(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Long) As Long
#Else
Public Function SQLite3BindInt32(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Long) As Long
#End If
    SQLite3BindInt32 = sqlite3_bind_int(stmtHandle, OneBasedParamIndex, Value)
End Function

#If WIN64 Then
Public Function SQLite3BindDate(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Date) As Long
#Else
Public Function SQLite3BindDate(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Date) As Long
#End If
    SQLite3BindDate = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, ToJulianDay(Value))
End Function

#If WIN64 Then
Public Function SQLite3BindBlob(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByRef Value() As Byte) As Long
#Else
Public Function SQLite3BindBlob(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByRef Value() As Byte) As Long
#End If
    Dim Length As Long
    Length = UBound(Value) - LBound(Value) + 1
    SQLite3BindBlob = sqlite3_bind_blob(stmtHandle, OneBasedParamIndex, VarPtr(Value(0)), Length, SQLITE_TRANSIENT)
End Function

#If WIN64 Then
Public Function SQLite3BindNull(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As Long
#Else
Public Function SQLite3BindNull(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As Long
#End If
    SQLite3BindNull = sqlite3_bind_null(stmtHandle, OneBasedParamIndex)
End Function

#If WIN64 Then
Public Function SQLite3BindParameterCount(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3BindParameterCount(ByVal stmtHandle As Long) As Long
#End If
    SQLite3BindParameterCount = sqlite3_bind_parameter_count(stmtHandle)
End Function

#If WIN64 Then
Public Function SQLite3BindParameterName(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As String
#Else
Public Function SQLite3BindParameterName(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As String
#End If
    SQLite3BindParameterName = UTFlib.Utf8PtrToString(sqlite3_bind_parameter_name(stmtHandle, OneBasedParamIndex))
End Function

#If WIN64 Then
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As LongPtr, ByVal paramName As String) As Long
#Else
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As Long, ByVal paramName As String) As Long
#End If
    Dim buf() As Byte
    buf = UTFlib.StringToUtf8Bytes(paramName)
    SQLite3BindParameterIndex = sqlite3_bind_parameter_index(stmtHandle, VarPtr(buf(0)))
End Function

#If WIN64 Then
Public Function SQLite3ClearBindings(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3ClearBindings(ByVal stmtHandle As Long) As Long
#End If
    SQLite3ClearBindings = sqlite3_clear_bindings(stmtHandle)
End Function


'=====================================================================================
' Backup
Public Function SQLite3Sleep(ByVal timeToSleepInMs As Long) As Long
    SQLite3Sleep = sqlite3_sleep(timeToSleepInMs)
End Function

#If WIN64 Then
Public Function SQLite3BackupInit(ByVal dbHandleDestination As LongPtr, ByVal destinationName As String, ByVal dbHandleSource As LongPtr, ByVal sourceName As String) As LongPtr
#Else
Public Function SQLite3BackupInit(ByVal dbHandleDestination As Long, ByVal destinationName As String, ByVal dbHandleSource As Long, ByVal sourceName As String) As Long
#End If
    Dim bufDestinationName() As Byte
    Dim bufSourceName() As Byte
    bufDestinationName = UTFlib.StringToUtf8Bytes(destinationName)
    bufSourceName = UTFlib.StringToUtf8Bytes(sourceName)
    SQLite3BackupInit = sqlite3_backup_init(dbHandleDestination, VarPtr(bufDestinationName(0)), dbHandleSource, VarPtr(bufSourceName(0)))
End Function

#If WIN64 Then
Public Function SQLite3BackupFinish(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupFinish(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupFinish = sqlite3_backup_finish(backupHandle)
End Function

#If WIN64 Then
Public Function SQLite3BackupStep(ByVal backupHandle As LongPtr, ByVal numberOfPages) As Long
#Else
Public Function SQLite3BackupStep(ByVal backupHandle As Long, ByVal numberOfPages) As Long
#End If
    SQLite3BackupStep = sqlite3_backup_step(backupHandle, numberOfPages)
End Function

#If WIN64 Then
Public Function SQLite3BackupPageCount(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupPageCount(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupPageCount = sqlite3_backup_pagecount(backupHandle)
End Function

#If WIN64 Then
Public Function SQLite3BackupRemaining(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupRemaining(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupRemaining = sqlite3_backup_remaining(backupHandle)
End Function


' Date Helpers
Public Function ToJulianDay(oleDate As Date) As Double
    Const JULIANDAY_OFFSET As Double = 2415018.5
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Public Function FromJulianDay(julianDay As Double) As Date
    Const JULIANDAY_OFFSET As Double = 2415018.5
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function
