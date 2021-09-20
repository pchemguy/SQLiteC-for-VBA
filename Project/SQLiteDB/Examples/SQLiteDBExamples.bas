Attribute VB_Name = "SQLiteDBExamples"
'@Folder "SQLiteDB.Examples"
'@IgnoreModule ProcedureNotUsed, VariableNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module
Option Compare Text

Private Const LIB_NAME As String = "SQLiteDBVBA"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = REL_PREFIX & "SQLiteDBVBA.db"
    Dim TargetDb As String
    TargetDb = REL_PREFIX & "Dest.db"

    '@Ignore FunctionReturnValueDiscarded
    SQLiteDB.CloneDb TargetDb, SourceDb
End Sub


Private Sub SetJournalMode()
    Dim FileName As String
    FileName = REL_PREFIX & "TestA.db"
    
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(FileName)
    DbManager.AttachDatabase REL_PREFIX & "TestB.db"
    DbManager.AttachDatabase REL_PREFIX & "TestC.db"
    
    DbManager.JournalModeSet "WAL", "ALL"
End Sub


Private Sub PrintTable()
    Dim OutputWS As Excel.Worksheet
    Set OutputWS = Buffer
        
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(FileName)
    
    Dim SQLTool As SQLlib
    Set SQLTool = SQLlib("contacts")
    SQLTool.Limit = 1000
    DbManager.DebugPrintRecordset SQLTool.SelectAll, OutputWS.Range("A1")
End Sub


Private Function GetFileBytes(ByVal FilePathName As String) As Byte()
    Dim FileHandle As Long
    Dim ReadBuffer(0 To 15) As Byte
    FileHandle = FreeFile
    Open FilePathName For Binary Access Read Write As FileHandle
    Get FileHandle, , ReadBuffer
    Close FileHandle
    GetFileBytes = ReadBuffer
End Function


Public Sub TestGetFileBytes()
    Dim FilePathName As String
    FilePathName = ThisWorkbook.Path & "\Library\SQLiteDBVBA\TestC.db"
    Dim ByteBuffer() As Byte
    ByteBuffer = GetFileBytes(FilePathName)
    Dim TextBuffer As String
    TextBuffer = StrConv(ByteBuffer, vbUnicode)
    Dim SQLiteDbSignature As String
    SQLiteDbSignature = "SQLite format 3" & vbNullChar
    Debug.Print SQLiteDbSignature = TextBuffer
End Sub


Public Function CheckAccessAndBasicIntegrity() As Boolean
    Dim FilePathName As String
    FilePathName = ThisWorkbook.Path & "\Library\SQLiteDBVBA\TestC.db"
    FilePathName = ThisWorkbook.Path & "\SQLiteDBVBA.db"
    FilePathName = ThisWorkbook.Path & "\Library\SQLiteDBVBA\TestC ACL.db"
    FilePathName = ThisWorkbook.Path & "\Library\SQLiteDBVBA\SQLiteDBVBA.db"
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim ErrNumber As Long
    Dim ErrSource As String
    Dim ErrDescription As String
    
    '''' ===== Checks for an existing or new database file ===== ''''
    
    '''' Get parent folder and verify it exists.
    Dim DbFilePath As String
    DbFilePath = fso.GetParentFolderName(FilePathName)
    On Error Resume Next
    Dim DbFolder As Scripting.Folder
    Set DbFolder = fso.GetFolder(DbFilePath)
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    '''' Negative result may also mean, e.g., ACL permission issues
    '''' The only expected error is PathNotFound. Possible reasons:
    ''''   - path format is illegal;
    ''''   - any path component does not exists;
    ''''   - any folder, except for the final is not accessible due to
    ''''     ACL permission settings.
    Select Case ErrNumber
        Case ErrNo.PathNotFoundErr
            Err.Raise ErrNumber, "SQLiteDB", "Path not found. " & _
                   "Check that path is legal, existent, and accessible (ACL)"
        Case Is <> 0
            Err.Raise ErrNumber, ErrSource, ErrDescription
    End Select
    
    '''' Path is OK.
    '''' Verify that folder is accessible - get file/subfolder count.
    On Error Resume Next
    Dim SubFolderCount As Long
    SubFolderCount = DbFolder.SubFolders.Count
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    '''' The only expected error is PermissionDenied due to ACL.
    Select Case ErrNumber
        Case ErrNo.PermissionDeniedErr
            Err.Raise ErrNumber, "SQLiteDB", "Access is denied to the folder" & _
                   " containing the database file. Check ACL permissions."
        Case Is <> 0
            Err.Raise ErrNumber, ErrSource, ErrDescription
    End Select
    
    '''' ===== Checks for an existing database file ===== ''''
    
    '''' Folder is accessible.
    '''' Verify that the file exists and its size is >=100 (SQLite header = 100).
    If Not fso.FileExists(FilePathName) Then
        Err.Raise ErrNo.FileNotFoundErr, "SQLiteDB", "Databse file not found"
    End If
    Dim DbFile As Scripting.File
    Set DbFile = fso.GetFile(FilePathName)
    If DbFile.Size < 100 Then
        Err.Raise ErrNo.AdoInvalidFileFormatErr, "SQLiteDB", "File is not " & _
                  "a database. SQLite header size is 100 bytes."
    End If
    
    '''' File size is OK.
    '''' Verify that the file is accessible.
    Const TRISTATE_OPEN_AS_ASCII As Long = TristateFalse
    Const TRISTATE_OPEN_AS_UNICODE As Long = TristateTrue
    On Error Resume Next
    Dim DbTextStream As Scripting.TextStream
    Set DbTextStream = DbFile.OpenAsTextStream(ForReading, TRISTATE_OPEN_AS_ASCII)
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    '''' The only expected error is PermissionDenied due to ACL.
    Select Case ErrNumber
        Case ErrNo.PermissionDeniedErr
            Err.Raise ErrNumber, "SQLiteDB", "Access denied to the " & _
                   "database file. Check ACL permissions and file locks."
        Case Is <> 0
            Err.Raise ErrNumber, ErrSource, ErrDescription
    End Select
    
    '''' File is accessible.
    '''' Verify that the database file is accessible for reading.
    Const MagicHeaderString As String = "SQLite format 3" & vbNullChar
    On Error Resume Next
    Dim FileSignature As String
    FileSignature = DbTextStream.Read(Len(MagicHeaderString))
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    '''' The only expected error is TextStreamReadErr: while a file stream can
    '''' be opened for reading, the file might still be locked. Apparently,
    '''' to test it, an actual read attempt must be made.
    Select Case ErrNumber
        Case ErrNo.TextStreamReadErr
            Err.Raise ErrNumber, "SQLiteDB", "Cannot read from the database " & _
                   "file. Most likely, the file is locked by another app."
        Case Is <> 0
            Err.Raise ErrNumber, ErrSource, ErrDescription
    End Select
        
    '''' Reading is successful.
    '''' Verify magic string.
    If Not FileSignature = MagicHeaderString Then
        Err.Raise ErrNo.AdoInvalidFileFormatErr, "SQLiteDB", "Database " & _
                  "file is damaged: the magic string did not match."
    End If
End Function


    
'   Set DbTextStream = fso.GetFile(FilePathName & "-shm").OpenAsTextStream(ForReading, TRISTATE_OPEN_AS_ASCII)
'   Dim Buffer As String
'   Buffer = DbTextStream.Read(4)
'
'    Debug.Print fso.FolderExists("F:\Archive\Business\FID\PolMaFID\Drafts\Knowledge Management System\VBA\SQLiteDB VBA Library\Library\SQLiteDBVBA\")
'
'    Debug.Print FileLen("F:\Archive\Business\FID\PolMaFID\Drafts\Knowledge Management System\VBA\SQLiteDB VBA Library\Library\SQLiteDBVBA\QQ")
    
    '''' Split FilePathName
'    Dim SplitPosition As Long
'    SplitPosition = InStrRev(FilePathName, Application.PathSeparator)
'    Dim FileName As String
'    FileName = Mid$(FilePathName, SplitPosition + 1)
'
'    Const MagicHeaderString As String = "SQLite format 3" & vbNullChar
'    Dim FileHandle As Long
'    Dim ReadBuffer(0 To Len(MagicHeaderString) - 1) As Byte
'
'    FileHandle = FreeFile
'    Open FilePathName For Binary Access Read Write As FileHandle
'    Get FileHandle, , ReadBuffer
'    Close FileHandle
'    VerifyMagicHeader = (MagicHeaderString = StrConv(ReadBuffer, vbUnicode))

