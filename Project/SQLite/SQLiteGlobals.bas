Attribute VB_Name = "SQLiteGlobals"
'@Folder "SQLite"
Option Explicit
Option Compare Text

Public Enum EnvArchEnum
    ENVARCH_NOTSUP = -1&
    ENVARCH_NATIVE = 1&
    ENVARCH_32ON64 = 2&
End Enum


'@Description "Determines environment type (native, X32onX64, or not supported)"
Public Function GetEnvX32X64Type() As EnvArchEnum
Attribute GetEnvX32X64Type.VB_Description = "Determines environment type (native, X32onX64, or not supported)"
    '@Ignore SelfAssignedDeclaration
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim OfficeArch As Long
    OfficeArch = Val(Right$(ARCH, 2))
    
    '''' Check actual Windows architecture. Environ returns a virtual value.
    On Error Resume Next
    Dim ProcArch As String
    ProcArch = wsh.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\" & _
        "Session Manager\Environment\PROCESSOR_ARCHITECTURE")
    On Error GoTo 0
    If Len(ProcArch) = 0 Then
        GetEnvX32X64Type = ENVARCH_NOTSUP
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

    If OfficeArch = WindowsArch And WindowsArch <> 0 Then
        GetEnvX32X64Type = ENVARCH_NATIVE
    ElseIf OfficeArch = 32 And WindowsArch = 64 Then
        GetEnvX32X64Type = ENVARCH_32ON64
    Else
        GetEnvX32X64Type = ENVARCH_NOTSUP
    End If
End Function
