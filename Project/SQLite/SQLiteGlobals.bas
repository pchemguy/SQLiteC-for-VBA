Attribute VB_Name = "SQLiteGlobals"
'@Folder "SQLite"
Option Explicit
Option Compare Text

#If Win64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
    Public Const vbLongLong As Long = 20&
#End If


Public Enum EnvArchEnum
    ENVARCH_NOTSUP = -1&
    ENVARCH_NATIVE = 1&
    ENVARCH_32ON64 = 2&
End Enum


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
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    
    Dim OfficeArch As Long
    OfficeArch = Val(Right(ARCH, 2))
    
    On Error Resume Next
    Dim ProcArch As String
    ProcArch = wsh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\" & _
        "Session Manager\Environment\PROCESSOR_ARCHITECTURE")
    On Error GoTo 0
    If Len(ProcArch) = 0 Then
        Debug.Print "Failed to determine Win/Office architecture or unsupported."
        SQLite3ODBCDriverCheck = False
        Exit Function
    End If
    
    Dim WindowsArch As Long
    Select Case ProcArch
        Case "AMD64"
            WindowsArch = 64
        Case "x86"
            WindowsArch = 32
        Case Else
            WindowsArch = 0
    End Select
    
    Dim RegPrefix As String
    Dim EnvArch As EnvArchEnum
    If OfficeArch = WindowsArch And WindowsArch <> 0 Then
        EnvArch = ENVARCH_NATIVE
        RegPrefix = "HKLM\SOFTWARE\ODBC\ODBC.INI\"
    ElseIf OfficeArch = 32 And WindowsArch = 64 Then
        EnvArch = ENVARCH_32ON64
        RegPrefix = "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\"
    Else
        EnvArch = ENVARCH_NOTSUP
        Debug.Print "Failed to determine Win/Office architecture or unsupported."
        SQLite3ODBCDriverCheck = False
        Exit Function
    End If
    
    Const SYSTEM_NATIVE As String = "System32"
    Const SYSTEM_32ON64 As String = "SysWOW64"
        
    Dim SQLite3ODBCDriverPath As String
    On Error Resume Next
        SQLite3ODBCDriverPath = _
            wsh.RegRead(RegPrefix & "SQLite3 Datasource\Driver")
    On Error GoTo 0
    If Len(SQLite3ODBCDriverPath) = 0 Then
        Debug.Print "Failed to verify SQLite3ODBC driver availability"
        Exit Function
    End If
    If EnvArch = ENVARCH_32ON64 Then
        SQLite3ODBCDriverPath = _
            Replace(SQLite3ODBCDriverPath, SYSTEM_NATIVE, SYSTEM_32ON64)
    End If
    
    If fso.FileExists(SQLite3ODBCDriverPath) Then
        Debug.Print "SQLite3ODBC driver appears to be available."
        SQLite3ODBCDriverCheck = True
    Else
        Debug.Print "Failed to verify SQLite3ODBC driver availability"
        SQLite3ODBCDriverCheck = False
    End If
End Function
