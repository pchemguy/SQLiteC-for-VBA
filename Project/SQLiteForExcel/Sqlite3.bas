Attribute VB_Name = "Sqlite3"
'@Folder "SQLiteForExcel"
'@IgnoreModule
Option Explicit

'=====================================================================================
' Statement column access (0-based indices)



#If Win64 Then
Public Function SQLite3ColumnText(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As String
#Else
Public Function SQLite3ColumnText(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As String
#End If
    SQLite3ColumnText = UTF8PtrToString(sqlite3_column_text(stmtHandle, ZeroBasedColIndex))
End Function

#If Win64 Then
Public Function SQLite3ColumnDate(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long) As Date
#Else
Public Function SQLite3ColumnDate(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long) As Date
#End If
    SQLite3ColumnDate = FromJulianDay(ColumnDouble(stmtHandle, ZeroBasedColIndex))
End Function

'=====================================================================================
' Statement bindings

#If Win64 Then
Public Function SQLite3BindDouble(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Double) As Long
#Else
Public Function SQLite3BindDouble(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Double) As Long
#End If
    SQLite3BindDouble = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, Value)
End Function

#If Win64 Then
Public Function SQLite3BindInt32(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Long) As Long
#Else
Public Function SQLite3BindInt32(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Long) As Long
#End If
    SQLite3BindInt32 = sqlite3_bind_int(stmtHandle, OneBasedParamIndex, Value)
End Function

#If Win64 Then
Public Function SQLite3BindDate(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByVal Value As Date) As Long
#Else
Public Function SQLite3BindDate(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByVal Value As Date) As Long
#End If
    SQLite3BindDate = sqlite3_bind_double(stmtHandle, OneBasedParamIndex, ToJulianDay(Value))
End Function

#If Win64 Then
Public Function SQLite3BindBlob(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long, ByRef Value() As Byte) As Long
#Else
Public Function SQLite3BindBlob(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long, ByRef Value() As Byte) As Long
#End If
    Dim length As Long
    length = UBound(Value) - LBound(Value) + 1
    SQLite3BindBlob = sqlite3_bind_blob(stmtHandle, OneBasedParamIndex, VarPtr(Value(0)), length, SQLITE_TRANSIENT)
End Function

#If Win64 Then
Public Function SQLite3BindNull(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As Long
#Else
Public Function SQLite3BindNull(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As Long
#End If
    SQLite3BindNull = sqlite3_bind_null(stmtHandle, OneBasedParamIndex)
End Function

#If Win64 Then
Public Function SQLite3BindParameterCount(ByVal stmtHandle As LongPtr) As Long
#Else
Public Function SQLite3BindParameterCount(ByVal stmtHandle As Long) As Long
#End If
    SQLite3BindParameterCount = sqlite3_bind_parameter_count(stmtHandle)
End Function

#If Win64 Then
Public Function SQLite3BindParameterName(ByVal stmtHandle As LongPtr, ByVal OneBasedParamIndex As Long) As String
#Else
Public Function SQLite3BindParameterName(ByVal stmtHandle As Long, ByVal OneBasedParamIndex As Long) As String
#End If
    SQLite3BindParameterName = UTF8PtrToString(sqlite3_bind_parameter_name(stmtHandle, OneBasedParamIndex))
End Function

#If Win64 Then
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As LongPtr, ByVal paramName As String) As Long
#Else
Public Function SQLite3BindParameterIndex(ByVal stmtHandle As Long, ByVal paramName As String) As Long
#End If
    Dim buf() As Byte
    buf = StringToUtf8Bytes(paramName)
    SQLite3BindParameterIndex = sqlite3_bind_parameter_index(stmtHandle, VarPtr(buf(0)))
End Function

#If Win64 Then
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

#If Win64 Then
Public Function SQLite3BackupInit(ByVal dbHandleDestination As LongPtr, ByVal destinationName As String, ByVal dbHandleSource As LongPtr, ByVal sourceName As String) As LongPtr
#Else
Public Function SQLite3BackupInit(ByVal dbHandleDestination As Long, ByVal destinationName As String, ByVal dbHandleSource As Long, ByVal sourceName As String) As Long
#End If
    Dim bufDestinationName() As Byte
    Dim bufSourceName() As Byte
    bufDestinationName = StringToUtf8Bytes(destinationName)
    bufSourceName = StringToUtf8Bytes(sourceName)
    SQLite3BackupInit = sqlite3_backup_init(dbHandleDestination, VarPtr(bufDestinationName(0)), dbHandleSource, VarPtr(bufSourceName(0)))
End Function

#If Win64 Then
Public Function SQLite3BackupFinish(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupFinish(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupFinish = sqlite3_backup_finish(backupHandle)
End Function

#If Win64 Then
Public Function SQLite3BackupStep(ByVal backupHandle As LongPtr, ByVal numberOfPages) As Long
#Else
Public Function SQLite3BackupStep(ByVal backupHandle As Long, ByVal numberOfPages) As Long
#End If
    SQLite3BackupStep = sqlite3_backup_step(backupHandle, numberOfPages)
End Function

#If Win64 Then
Public Function SQLite3BackupPageCount(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupPageCount(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupPageCount = sqlite3_backup_pagecount(backupHandle)
End Function

#If Win64 Then
Public Function SQLite3BackupRemaining(ByVal backupHandle As LongPtr) As Long
#Else
Public Function SQLite3BackupRemaining(ByVal backupHandle As Long) As Long
#End If
    SQLite3BackupRemaining = sqlite3_backup_remaining(backupHandle)
End Function

