Attribute VB_Name = "ShellRoutines"
'@Folder "Common.Shared"
'@IgnoreModule ConstantNotUsed, AssignmentNotUsed, VariableNotUsed, ProcedureNotUsed, UseMeaningfulName
Option Explicit

' The WaitForSingleObject function returns when one of the following occurs:
' - The specified object is in the signaled state.
' - The time-out interval elapses.
'
' The dwMilliseconds parameter specifies the time-out interval, in milliseconds.
' The function returns if the interval elapses, even if the object’s state is
' nonsignaled. If dwMilliseconds is zero, the function tests the object’s state
' and returns immediately. If dwMilliseconds is INFINITE, the function’s time-out
' interval never elapses.
'
' This example waits an INFINITE amount of time for the process to end. As a
' result this process will be frozen until the shelled process terminates. The
' down side is that if the shelled process hangs, so will this one.
'
' A better approach is to wait a specific amount of time. Once the time-out
' interval expires, test the return value. If it is WAIT_TIMEOUT, the process
' is still not signaled. Then you can either wait again or continue with your
' processing.
'
' DOS Applications:
' Waiting for a DOS application is tricky because the DOS window never goes
' away when the application is done. To get around this, prefix the app that
' you are shelling to with "command.com /c".
'
' For example: lPid = Shell("command.com /c " & txtApp.Text, vbNormalFocus)
'
' To get the return code from the DOS app, see the attached text file.
'

Public Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Public Const WS_THICKFRAME As Long = &H40000     'Style to add a sizable frame
Public Const SW_SHOW As Long = 5

#If VBA7 Then
    'For 64-Bit versions of Excel
    
    '''' Window functions
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    
    '''' Process functions
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    'For 32-Bit versions of Excel
    
    '''' Window functions
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
    
    '''' Process functions
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Const SYNCHRONIZE As Long = &H100000
'
' Wait forever
Const INFINITE As Long = &HFFFF
'
' The state of the specified object is signaled
Const WAIT_OBJECT_0 As Long = 0
'
' The time-out interval elapsed & the object’s state is not signaled
Const WAIT_TIMEOUT As Long = 1000


Public Function ReadLines(ByVal FilePath As String) As Variant
    Dim Handle As Long
    Handle = FreeFile
    Open FilePath For Input As Handle
    
    Dim Buffer As String
    Buffer = Input$(LOF(Handle), Handle)
    If Right$(Buffer, Len(vbNewLine)) = vbNewLine Then
        Buffer = Left$(Buffer, Len(Buffer) - Len(vbNewLine))
    End If
    ReadLines = Split(Buffer, vbNewLine)
    Close Handle
End Function


' Runs shell command, waits for completions, sends stdout to file and returns stdout as an array of strings
Public Function SyncRun(ByVal Command As String, Optional ByVal redirectStdout As Boolean = True) As Variant
    Dim cli As String
    If redirectStdout Then
        Dim GUID As String
        GUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
        Dim sys As IWshRuntimeLibrary.WshShell
        Set sys = New IWshRuntimeLibrary.WshShell
        Dim TempFile As String
        TempFile = sys.ExpandEnvironmentStrings("%temp%\stdout-") & GUID & ".txt"
        cli = Command & " >""" & TempFile & """"
    Else
        cli = Command
    End If

    Dim PID As Long
    PID = Shell(cli, vbHide)

    If PID <> 0 Then
        'Get a handle to the shelled process.
        #If VBA7 Then
            Dim Handle As LongPtr
            Handle = OpenProcess(SYNCHRONIZE, 0, PID)
        #Else
            Dim Handle As Long
            Handle = OpenProcess(SYNCHRONIZE, 0, PID)
        #End If
        
        'If successful, wait for the application to end and close the handle.
        If Handle <> 0 Then
            Dim Result As Long
            Result = WaitForSingleObject(Handle, WAIT_TIMEOUT)
            CloseHandle hObject:=Handle
        End If
    End If
    If redirectStdout Then
        SyncRun = ReadLines(TempFile)
    End If
End Function


'@Ignore ProcedureNotUsed
Private Sub Test()
    Dim cmdline As String
    Dim output As Variant
    
    cmdline = "cmd /c dir /b c:\windows"
    output = SyncRun(cmdline)
    'ShellSync ("cmd /c dir /b c:\windows |clip")
End Sub
