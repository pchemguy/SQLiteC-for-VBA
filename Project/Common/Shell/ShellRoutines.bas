Attribute VB_Name = "ShellRoutines"
'@Folder "Common.Shell"
'@IgnoreModule ConstantNotUsed, AssignmentNotUsed, VariableNotUsed, UseMeaningfulName, ProcedureNotUsed
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


'''' Reads lines from the text file specified by FilePath and returns an array
'''' of strings representing the file. FilePath must be validated by the caller
'''' if necessary.
''''
'@Description "Reads lines from a text file"
Public Function ReadLines(ByVal FilePath As String) As Variant
Attribute ReadLines.VB_Description = "Reads lines from a text file"
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


'''' Uses WinAPI to run provided "Command" synchronously (waits for process exit)
'''' and optionally (by default) returns the output as an array of strings
''''
'@Description "Runs shell command synchronously and returns stdout"
Public Function SyncRun(ByVal Command As String, Optional ByVal RedirectStdout As Boolean = True) As Variant
Attribute SyncRun.VB_Description = "Runs shell command synchronously and returns stdout"
    Dim cli As String
    If RedirectStdout Then
        Dim GUID As String
        GUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
        Dim WshLib As IWshRuntimeLibrary.wshShell
        Set WshLib = New IWshRuntimeLibrary.wshShell
        Dim TempFile As String: TempFile = WshLib.ExpandEnvironmentStrings("%temp%\stdout-") & GUID & ".txt"
        cli = Command & " >""" & TempFile & """"
    Else
        cli = Command
    End If

    Dim pid As Long
    pid = Shell(cli, vbHide)

    If pid <> 0 Then
        'Get a handle to the shelled process.
        #If VBA7 Then
            Dim Handle As LongPtr: Handle = OpenProcess(SYNCHRONIZE, 0, pid)
        #Else
            Dim Handle As Long: Handle = OpenProcess(SYNCHRONIZE, 0, pid)
        #End If
        
        'If successful, wait for the application to end and close the handle.
        If Handle <> 0 Then
            Dim Result As Long: Result = WaitForSingleObject(Handle, WAIT_TIMEOUT)
            CloseHandle hObject:=Handle
        End If
    End If
    If RedirectStdout Then
        SyncRun = ReadLines(TempFile)
    End If
End Function


'@Description "Returns an array of file names matching the glob in the given directory"
Public Function DirList(ByVal Path As String, Optional ByVal FileMask As String = "*.*") As Variant
Attribute DirList.VB_Description = "Returns an array of file names matching the glob in the given directory"
    DirList = SyncRun("cmd /c dir /b """ & Path & FileMask & """")
End Function


'''' Attempts to determine full pathname for provided argument
''''
'''' Takes a file name with optional relative or absolute path and checks if
'''' file exists (relative to the default current directory, which should be
'''' <Thisworkbook.Path>).
'''' If Not found And DefaultExts Is Empty Then either an error is raised
'''' (AllowNonExistent is False) or FilePathName (with ThisWorkbook.Path prefix
'''' if FilePathName contains no PathSeparator) is returned (AllowNonExistent
'''' is True).
'''' Otherwise, if the second argument is provided, the current directory is
'''' searched for file with the name Thisworkbook.Name (without extension). The
'''' items from the second argument are checked sequentially as extension
'''' candidates. If file is found, the search is stopped. If the list of
'''' extension candidates is exhausted, an error is raised.
''''
'''' VerifyOrGetDefaultPath should always be called with two arguments. If
'''' the second argument is not Empty, the third argument will be ignored.
'''' The first argument is allowed to refer to a non-existent file only
'''' if the second argument is Empty and the third argument is True.
''''
'''' Args:
''''   FilePathName (string):
''''     File name with extension. Relative or absolute path name may also be provided.
''''   DefaultExts (array of strings, optional, Empty):
''''     Array containing extensions to be used for the search.
''''   AllowNonExistent (boolean, optional, False):
''''     If set to True, FilePathName may point to a non-existent file.
''''
'''' Returns:
''''   String FilePathName to the first found file
''''
'''' Raises:
''''   ErrNo.FileNotFoundErr:
''''     if FilePathName cannot be resolved as a valid PathName and default search fails as well.
''''
'''' Examples:
''''   Raises error:
''''     >>> ?VerifyOrGetDefaultPath("")
''''     Raises "FileNotFoundErr" error
''''
''''     TODO: Add unit tests
''''     >>> ?VerifyOrGetDefaultPath("___.___")
''''     Raises "FileNotFoundErr" error
''''
''''     TODO: Add unit tests
''''     >>> ?VerifyOrGetDefaultPath("___.___", Array("___"), True)
''''     Raises "FileNotFoundErr" error
''''
''''     >>> ?VerifyOrGetDefaultPath("", , True)
''''     Raises "FileNotFoundErr" error
''''
''''   Call with a file name:
''''     >>> ?VerifyOrGetDefaultPath("SQLiteDB.db")
''''     "<Thisworkbook.Path>\SQLiteDB.db"
''''
''''   Call with specified extensions:
''''     >>> ?VerifyOrGetDefaultPath("", Array("sqlite", "db"))
''''     "<Thisworkbook.Path>\SQLiteDB.db"
''''
''''   TODO: Add unit tests
''''   Allow non-existent:
''''     >>> ?VerifyOrGetDefaultPath("___.___", , True)
''''     "<Thisworkbook.Path>\___.___"
''''
''''     >>> ?VerifyOrGetDefaultPath(Application.PathSeparator & "___.___", , True)
''''     Application.PathSeparator & "___.___"
''''
'@Description "Attempts to determine full pathname for provided argument"
Public Function VerifyOrGetDefaultPath(ByVal FilePathName As String, _
                              Optional ByVal DefaultExts As Variant = Empty, _
                              Optional ByVal AllowNonExistent As Boolean = False) As String
Attribute VerifyOrGetDefaultPath.VB_Description = "Attempts to determine full pathname for provided argument"
    '''' Check if FilePathName is a valid path to an existing file.
    '''' If yes, return it.
    On Error Resume Next
    Dim FileExist As Variant
    If Len(FilePathName) > 0 Then FileExist = Dir$(FilePathName)
    On Error GoTo 0
    If Len(FileExist) > 0 Then
        VerifyOrGetDefaultPath = FilePathName
        Exit Function
    End If
    
    '''' Check if supplied name is a file in ThisWorkbook.Path
    On Error Resume Next
    FileExist = Dir$(ThisWorkbook.Path & Application.PathSeparator & FilePathName)
    On Error GoTo 0
    If FileExist = FilePathName Then
        VerifyOrGetDefaultPath = ThisWorkbook.Path & Application.PathSeparator & FilePathName
        Exit Function
    End If
    
    If IsEmpty(DefaultExts) Then
        If AllowNonExistent And Len(FilePathName) > 0 Then
            If InStr(FilePathName, Application.PathSeparator) Then
                VerifyOrGetDefaultPath = FilePathName
                Exit Function
            Else
                VerifyOrGetDefaultPath = ThisWorkbook.Path & Application.PathSeparator & FilePathName
                Exit Function
            End If
        Else
            VBA.Err.Raise Number:=ErrNo.FileNotFoundErr, Source:="SQLiteDB", _
                          Description:="File <" & FilePathName & "> not found!"
        End If
    End If
    
    '''' Check defaults:
    ''''   - path: ThisWorkbook.Path
    ''''   - name: ThisWorkbook.Name (without extension)
    ''''   - exts: DefaultExts
    Dim DefaultName As String: DefaultName = ThisWorkbook.Name
    Dim DotPos As Long: DotPos = InStr(Len(DefaultName) - 5, DefaultName, ".xl", vbTextCompare)
    DefaultName = Left$(DefaultName, DotPos)
    Dim DefaultPath As String
    DefaultPath = ThisWorkbook.Path & Application.PathSeparator & DefaultName
    
    Dim ExtIndex As Long
    Dim CheckedPath As String
    For ExtIndex = LBound(DefaultExts) To UBound(DefaultExts)
        CheckedPath = DefaultPath & DefaultExts(ExtIndex)
        FileExist = Dir$(CheckedPath)
        If Len(FileExist) > 0 Then
            VerifyOrGetDefaultPath = CheckedPath
            Exit Function
        End If
    Next ExtIndex
    
    If Len(FileExist) = 0 Then
        VBA.Err.Raise Number:=ErrNo.FileNotFoundErr, Source:="SQLiteDB", Description:="File <" & FilePathName & "> not found!"
    End If
End Function


Private Sub Test()
    Dim cmdline As String
    Dim Output As Variant
    
    cmdline = "cmd /c dir /b c:\windows"
    Output = SyncRun(cmdline)
    'ShellSync ("cmd /c dir /b c:\windows |clip")
End Sub
