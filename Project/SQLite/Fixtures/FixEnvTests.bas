Attribute VB_Name = "FixEnvTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext, SelfAssignedDeclaration
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "ODBCCP32" ( _
    ByVal lpszBuf As String, ByVal cbBufMax As Long, ByRef pcbBufOut As Long) As Long
#Else
Private Declare Function SQLGetInstalledDrivers Lib "ODBCCP32" ( _
    ByVal lpszBuf As String, ByVal cbBufMax As Long, ByRef pcbBufOut As Long) As Long
#End If
          
Private Const MODULE_NAME As String = "FixEnvTests"
Private TestCounter As Long

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
End Sub


'===================================================='
'=================== TEST FIXTURES =================='
'===================================================='


'''' This function attempts to confirm that the standard registry key for the
'''' SQLite3ODBC driver is present and that the file driver exists. No attempt
'''' is made to verify its usability.
''''
'''' Attempt to determine environment (native X32onX32 or X64onX64) or X32onX64.
'''' If successfull, try retrieving SQLite3ODBC driver file pathname from the
'''' standard registry key (adjusted to the type of environment, if necessary).
'''' If successful, adjust path to the type of environment, if necessary, and
'''' check if file driver exists. If successful, return true, or false otherwise.
''''
'@Description "Checks if SQLite3ODBC diver is available."
Public Function SQLite3ODBCDriverCheck() As Boolean
Attribute SQLite3ODBCDriverCheck.VB_Description = "Checks if SQLite3ODBC diver is available."
    Const SQLITE3_ODBC_NAME As String = "SQLite3 ODBC Driver"
    
    '''' Check if SQLGetInstalledDrivers contains the standard SQLite3ODBC driver
    '''' description. Fail if not found.
    Dim Buffer As String
    Buffer = String(2000, vbNullChar)
    Dim ActualSize As Long: ActualSize = 0 '''' RD ByRef workaround
    Dim Result As Boolean
    Result = SQLGetInstalledDrivers(Buffer, Len(Buffer) * 2, ActualSize)
    Result = InStr(Replace(Left$(Buffer, ActualSize - 1), vbNullChar, vbLf), _
                   SQLITE3_ODBC_NAME)
    If Not Result Then GoTo DRIVER_NOT_FOUND:
    
    Dim ODBCINSTPrefix As String
    Dim EnvArch As EnvArchEnum
    EnvArch = GetEnvX32X64Type()
    Select Case EnvArch
        Case ENVARCH_NATIVE
            ODBCINSTPrefix = "HKLM\SOFTWARE\ODBC\ODBCINST.INI\"
        Case ENVARCH_32ON64
            ODBCINSTPrefix = "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\"
        Case ENVARCH_NOTSUP
            Logger.Logg "Failed to determine Win/Office architecture or " & _
                        "unsupported.", , DEBUGLEVEL_ERROR
            SQLite3ODBCDriverCheck = False
            Exit Function
    End Select
    
    '''' Query standard ODBCINST.INI registry keys
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim SQLite3ODBCDriverInstalled As Boolean
    Dim RegPath As String
    RegPath = ODBCINSTPrefix & "ODBC Drivers\" & SQLITE3_ODBC_NAME
    SQLite3ODBCDriverInstalled = False
    On Error Resume Next
        SQLite3ODBCDriverInstalled = (wsh.RegRead(RegPath) = "Installed")
        If Not SQLite3ODBCDriverInstalled Then GoTo DRIVER_NOT_FOUND:
    On Error GoTo 0
    SQLite3ODBCDriverInstalled = False
    RegPath = ODBCINSTPrefix & SQLITE3_ODBC_NAME & "\Driver"
    Dim SQLite3ODBCDriverPath As String
    On Error Resume Next
        SQLite3ODBCDriverPath = wsh.RegRead(RegPath)
        If Len(SQLite3ODBCDriverPath) = 0 Then GoTo DRIVER_NOT_FOUND:
    On Error GoTo 0
    Const SYSTEM_NATIVE As String = "System32"
    Const SYSTEM_32ON64 As String = "SysWOW64"
    If EnvArch = ENVARCH_32ON64 Then
        SQLite3ODBCDriverPath = _
            Replace(SQLite3ODBCDriverPath, SYSTEM_NATIVE, SYSTEM_32ON64)
    End If
    
    '''' Check if driver file exists
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    If Not fso.FileExists(SQLite3ODBCDriverPath) Then GoTo DRIVER_NOT_FOUND:
    
    Logger.Logg "SQLite3ODBC driver appears to be available.", , DEBUGLEVEL_INFO
    SQLite3ODBCDriverCheck = True
    Exit Function
    
DRIVER_NOT_FOUND:
    Logger.Logg "Failed to verify SQLite3ODBC driver availability", , DEBUGLEVEL_ERROR
    SQLite3ODBCDriverCheck = False
    Exit Function
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Environment")
Private Sub ztc_VerifiesSQLiteODBCDriver()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
Assert:
    Assert.IsTrue SQLite3ODBCDriverCheck(), "SQLite3ODBC driver not found."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
