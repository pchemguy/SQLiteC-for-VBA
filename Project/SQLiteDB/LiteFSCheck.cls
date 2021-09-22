VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteFSCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides access and integrity check functionality"
'@Folder "SQLiteDB"
'@ModuleDescription "Provides access and integrity check functionality"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
'@IgnoreModule ValueRequired: false positive with ADODB memebers
'@IgnoreModule SelfAssignedDeclaration: ignore for FileSystemObject
''''
'''' This module incorporates OS based access and basic integrity checks of
'''' the database file. This module also provides extended integrity checks
'''' via the SQLite engine. The latter checks should be performed using a
'''' dedicated ADODB Connection object with the target database being the
'''' attached as the "main" and only database. (While those checks should
'''' should run on specific alias, I could not get it to work properly.)
'''' All these tests can be run on a transaction locked database, and they
'''' should introduce any locks.
''''
'''' OS-based checks verify that the provided database path is valid and points
'''' to an existing file. ACL access rights are also verified (both the parent
'''' folder and the database file ACL permissions are checked). Read access to
'''' the database file and its size (must be no smalller that the size of the
'''' SQLite header) are verified. Finally the SQLite magic string is verified.
''''
'''' OS-based checks are performed by three routines, ExistsAccesibleValid (the
'''' main entry), PathExistsAccessible, and FileAccessibleValid.
'''' PathExistsAccessible verifies that the file path is valid and that the
'''' parent folder is accessible. Its result is "the path is usable".
'''' FileAccessibleValid performs remaining OS-based file checks.
'''' ExistsAccesibleValid manages this part of tests. After a special case
'''' check, it probes "FileExists" test on the given path. If successful,
'''' PathExistsAccessible can be skipped. Otherwise, PathExistsAccessible
'''' will fail and is executed to narrow down the cause of the issue.
''''
'''' Engine-based checks are performed by IntegrityADODB. SQLite provides two
'''' pragmas for consistency checks, one performs a series of checks, except
'''' for the foreign key constraints checks, and the other one covers foreing
'''' keys.
''''
'''' Transaction-related lock checks should be performed before starting a
'''' a transaction or attempting a modification with auto transaction. These
'''' should use the primary Connection object and for this reason should not
'''' be a part of this module.
''''
Option Explicit

Private Type TLiteFSCheck
    MainDB As String '''' Absolute file pathname to the main database.
    Database As String
End Type
Private this As TLiteFSCheck


'''' Args:
''''   Database (string):
''''     Absolute path name to the database. No path processing should be done
''''       here. E.g., relative path will result in the PathNotFound error.
'@DefaultMember
Public Function Create(ByVal Database As String) As LiteFSCheck
    Dim Instance As LiteFSCheck
    Set Instance = New LiteFSCheck
    Instance.Init Database
    Set Create = Instance
End Function


Friend Sub Init(ByVal Database As String)
    this.Database = Database
End Sub


'@Description "Checks that provided path exists and is accessible."
Friend Function PathExistsAccessible(ByVal FilePathName As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
        
    '''' Get parent folder and verify it exists.
    Dim DbFilePath As String
    DbFilePath = fso.GetParentFolderName(FilePathName)
    If Len(DbFilePath) = 0 Then
        Err.Raise ErrNo.PathNotFoundErr, "LiteCheck", "Path is not absolute."
    End If
    
    On Error GoTo PATH_ERROR
        '''' Expected error: PathNotFound. Possible reasons:
        ''''   - path format is illegal;
        ''''   - any path component does not exists;
        ''''   - any folder, except for the final is not accessible due to
        ''''     ACL permission settings.
        Dim DbFolder As Scripting.Folder
        Set DbFolder = fso.GetFolder(DbFilePath)
        
        '''' Path is OK.
        '''' Verify that folder is accessible - get file/subfolder count.
        '''' Expected error: PermissionDenied due to ACL.
        '@Ignore VariableNotUsed: .Count property is accessed to test DbFolder access
        Dim SubFolderCount As Long
        SubFolderCount = DbFolder.SubFolders.Count
    On Error GoTo 0
        
    '''' Parent folder is OK.
    '''' Verify that the file exists.
    If Not fso.FileExists(FilePathName) Then
        Err.Raise ErrNo.FileNotFoundErr, "LiteCheck", "Databse file not found"
    End If
    
    PathExistsAccessible = True
    Exit Function
    
PATH_ERROR:
    Select Case Err.Number
        Case ErrNo.PathNotFoundErr
            Err.Raise Err.Number, "LiteCheck", "Path not found. " & _
                   "Check that path is legal, existent, and accessible (ACL)"
        Case ErrNo.PermissionDeniedErr
            Err.Raise Err.Number, "LiteCheck", "Access is denied to the folder" & _
                   " containing the database file. Check ACL permissions."
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & _
                      "Unexpected PathExistsAccessible error."
    End Select
End Function


'@Description "Verifies that the database file is accessible and valid."
Public Function FileAccessibleValid(ByVal FilePathName As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    
    '''' Verify that the file size is >=100 (SQLite header size is 100 bytes).
    Dim DbFile As Scripting.File
    Set DbFile = fso.GetFile(FilePathName)
    If DbFile.Size < 100 Then
        Err.Raise ErrNo.OLE_DB_ODBC_Err, "LiteCheck", "File is not " & _
                  "a database. SQLite header size is 100 bytes."
    End If
    
    On Error GoTo FILE_ERROR
        '''' File size is OK.
        '''' Verify that the file is accessible.
        Const TRISTATE_OPEN_AS_ASCII As Long = TristateFalse
        '@Ignore ConstantNotUsed
        Const TRISTATE_OPEN_AS_UNICODE As Long = TristateTrue
        '''' Expected error: PermissionDenied due to ACL.
        Dim DbTextStream As Scripting.TextStream
        Set DbTextStream = DbFile.OpenAsTextStream(ForReading, TRISTATE_OPEN_AS_ASCII)
        
        '''' File is accessible.
        '''' Verify that the database file is accessible for reading.
        Const MagicHeaderString As String = "SQLite format 3" & vbNullChar
        '''' Expected error: TextStreamReadErr due to another app's lock.
        Dim FileSignature As String
        FileSignature = DbTextStream.Read(Len(MagicHeaderString))
    On Error GoTo 0
        
    '''' Reading is successful.
    '''' Verify magic string.
    If Not FileSignature = MagicHeaderString Then
        Err.Raise ErrNo.OLE_DB_ODBC_Err, "LiteCheck", "Database " & _
                  "file is damaged: the magic string did not match."
    End If
    
    FileAccessibleValid = True
    Exit Function
    
FILE_ERROR:
    Select Case Err.Number
        Case ErrNo.PermissionDeniedErr
            Err.Raise Err.Number, "LiteCheck", "Access denied to the " & _
                   "database file. Check ACL permissions and file locks."
        Case ErrNo.TextStreamReadErr
            Err.Raise Err.Number, "LiteCheck", "Cannot read from the database " & _
                   "file. Most likely, the file is locked by another app."
        Case Is <> 0
            Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & _
                      "Unexpected FileExistsAccessibleValid error."
    End Select
End Function


'''' File access tests and basic integrity validation via FS OS functions.
'@Description "Runs SQLite database access and basic integrity checks via OS."
Friend Function ExistsAccesibleValid(ByVal FilePathName As String) As Boolean
    If this.Database = ":memory:" Then
        ExistsAccesibleValid = True
        Exit Function
    End If
    
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(this.Database) Then
        ExistsAccesibleValid = FileAccessibleValid(FilePathName)
    Else
        ExistsAccesibleValid = PathExistsAccessible(FilePathName)
    End If
End Function


