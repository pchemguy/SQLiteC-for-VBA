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
'@Folder "SQLite.Checks"
'@ModuleDescription "Provides access and integrity check functionality"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
'@IgnoreModule ValueRequired: false positive with ADODB memebers
'@IgnoreModule SelfAssignedDeclaration: ignore for FileSystemObject
''''
'''' This module incorporates OS based access and basic integrity checks of
'''' the database file. They verify that the provided database path is valid
'''' and points to an existing file. ACL access rights are also verified (both
'''' the parent folder and the database file ACL permissions are checked). Read
'''' access to the database file and its size (must be no smalller that the size
'''' of the SQLite header) are verified. Finally the SQLite magic string is
'''' verified.
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
Option Explicit

Private Type TLiteFSCheck
    DbPathName As String '''' Absolute file pathname to the main database.
    ErrNumber As Long
    ErrSource As String
    ErrDescription As String
    ErrStack As String
End Type
Private this As TLiteFSCheck


'''' Args:
''''   Database (string):
''''     Absolute path name to the database.
'@DefaultMember
Public Function Create(ByVal DbPathName As String, _
              Optional ByVal AllowNonExistent As Variant = Empty) As LiteFSCheck
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteFSCheck
    Set Instance = New LiteFSCheck
    Instance.Init DbPathName, AllowNonExistent
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbPathName As String, _
       Optional ByVal AllowNonExistent As Variant = Empty)
    this.DbPathName = ExistsAccesibleValid(DbPathName, AllowNonExistent)
End Sub

Public Property Get DatabasePathName() As String
    DatabasePathName = this.DbPathName
End Property

Public Property Get ErrNumber() As Long
    ErrNumber = this.ErrNumber
End Property

Public Property Get ErrSource() As String
    ErrSource = this.ErrSource
End Property

Public Property Get ErrDescription() As String
    ErrDescription = this.ErrDescription
End Property

Public Property Get ErrStack() As String
    ErrStack = this.ErrStack
End Property


'@Description "Checks that provided path exists and is accessible."
Friend Function PathExistsAccessible(ByVal FilePathName As String) As Boolean
Attribute PathExistsAccessible.VB_Description = "Checks that provided path exists and is accessible."
    Const ThisProcName As String = "PathExistsAccessible"
    Dim fso As New Scripting.FileSystemObject
        
    On Error GoTo PATH_ERROR
        '''' Get parent folder and verify it exists.
        Dim DbFilePath As String
        DbFilePath = fso.GetParentFolderName(FilePathName)
        If Not fso.FolderExists(DbFilePath) Then
            '''' Expected error: PathNotFound - Path is not absolute
            this.ErrDescription = "Database path (folder) is not found. " & _
                                  "Expected absolute path. Check ACL settings. " & _
                                  "Enable path resolution feature, if necessary."
            Err.Raise ErrNo.PathNotFoundErr, "LiteFSCheck", this.ErrDescription
        End If
        Dim DbFolder As Scripting.Folder
        Set DbFolder = fso.GetFolder(DbFilePath)
        
        '''' Path is OK.
        '''' Verify that folder is accessible - get file/subfolder count.
        '''' Expected error: PermissionDenied due to ACL.
        '@Ignore VariableNotUsed: .Count property is accessed to test DbFolder access
        Dim SubFolderCount As Long
        '@Ignore AssignmentNotUsed: .Count property is accessed to test DbFolder access
        SubFolderCount = DbFolder.SubFolders.Count
    On Error GoTo 0
    
    PathExistsAccessible = True
    Exit Function
    
PATH_ERROR:
    Select Case Err.Number
        Case ErrNo.PathNotFoundErr
            this.ErrDescription = Err.Description
        Case ErrNo.PermissionDeniedErr
            this.ErrDescription = Err.Description & vbNewLine & _
                                  "Access is denied to the folder containing " & _
                                  "the database file. Check ACL permissions."
        Case Else
            this.ErrDescription = Err.Description & vbNewLine & "Unexpected error."
    End Select

    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    Err.Raise Err.Number, "LiteFSCheck", _
              this.ErrDescription & vbNewLine & "Source: " & FilePathName
End Function


'@Description "Verifies that the database file is accessible and valid."
Public Function FileAccessibleValid(ByVal FilePathName As String) As String
Attribute FileAccessibleValid.VB_Description = "Verifies that the database file is accessible and valid."
    Const ThisProcName As String = "FileAccessibleValid"
    Dim fso As New Scripting.FileSystemObject
    
    On Error GoTo FILE_ERROR
        '''' Verify that the file exists.
        If Not fso.FileExists(FilePathName) Then
            this.ErrDescription = "Databse file is not found in the specified folder."
            Err.Raise ErrNo.FileNotFoundErr, "LiteFSCheck", this.ErrDescription
        End If
        
        '''' File exists
        '''' Verify that the file size is >=100 (SQLite header size is 100 bytes).
        Dim DbFile As Scripting.File
        Set DbFile = fso.GetFile(FilePathName)
        '''' Expected error: PermissionDenied due to ACL.
        If DbFile.Size < 100 Then
            this.ErrDescription = "File is not a database. " & _
                                  "SQLite header size is 100 bytes."
            Err.Raise ErrNo.OLE_DB_ODBC_Err, "LiteFSCheck", this.ErrDescription
        End If
            
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
    
        '''' Reading is successful.
        '''' Verify magic string.
        If Not FileSignature = MagicHeaderString Then
            this.ErrDescription = "Database file is damaged. " & _
                                  "The magic string did not match."
            Err.Raise ErrNo.OLE_DB_ODBC_Err, "LiteFSCheck", this.ErrDescription
        End If
    On Error GoTo 0
            
    FileAccessibleValid = FilePathName
    Exit Function
    
FILE_ERROR:
    Select Case Err.Number
        Case ErrNo.FileNotFoundErr
            this.ErrDescription = Err.Description
        Case ErrNo.PermissionDeniedErr
            this.ErrDescription = Err.Description & vbNewLine & _
                                  "Access is denied to the database file. " & _
                                  "Check ACL permissions and file locks."
        Case ErrNo.TextStreamReadErr
            this.ErrDescription = Err.Description & vbNewLine & _
                                  "Cannot read from the database file. " & _
                                  "The file might be locked by another app."
        Case ErrNo.OLE_DB_ODBC_Err
            this.ErrDescription = Err.Description
        Case Else
            this.ErrDescription = Err.Description & vbNewLine & "Unexpected error."
    End Select
    
    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    Err.Raise Err.Number, "LiteFSCheck", _
              this.ErrDescription & vbNewLine & "Source: " & FilePathName
End Function


'''' Args:
''''   DbPathName (string):
''''     Target database file pathname
''''
''''   AllowNonExistent (variant, optional, Empty):
''''     Controls additional path preprocessing and creatability
''''       - Empty        - new db file and path resolution are disabled
''''       - vbNullString - new db file only is enabled
''''       - True/False   - new db file is based on the value
''''                                    and path resolution is enabled
''''
'''' File access tests and basic integrity validation via FS/OS functions.
'@Description "Runs SQLite database access and basic integrity checks via OS."
Friend Function ExistsAccesibleValid(ByVal DbPathName As String, _
            Optional ByVal AllowNonExistent As Variant = Empty) As String
Attribute ExistsAccesibleValid.VB_Description = "Runs SQLite database access and basic integrity checks via OS."
    Const ThisProcName As String = "ExistsAccesibleValid"
    
    Dim FilePathName As String
    FilePathName = LCase$(DbPathName)
    '''' Check for special names
    Select Case FilePathName
        Case ":memory:"
            ExistsAccesibleValid = FilePathName
            Exit Function
        Case ":temp:", ":tmp:"
            ExistsAccesibleValid = Environ$("TEMP") & Application.PathSeparator & _
                Format$(Now, "yyyy_mm_dd-hh_mm_ss-") & Left$(GenerateGUID, 8) & ".db"
            Exit Function
        Case ":blank:"
            ExistsAccesibleValid = vbNullString
            Exit Function
    End Select
    
    Dim EnableResolve As Boolean
    EnableResolve = (VarType(AllowNonExistent) = vbBoolean)
    
    On Error GoTo FILEPATH_ERROR
    If EnableResolve Then
        '''' Expected error: FileNotFound
        FilePathName = VerifyOrGetDefaultPath( _
            DbPathName, Array("db", "sqlite"), AllowNonExistent)
    Else
        FilePathName = DbPathName
    End If
    
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(FilePathName) Then
        ExistsAccesibleValid = FileAccessibleValid(FilePathName)
    Else
        '@Ignore VariableNotUsed
        Dim DbPathNameOK As Boolean
        DbPathNameOK = PathExistsAccessible(FilePathName)
        '''' If PathExistsAccessible does not raise an error, Path is OK
        If IsEmpty(AllowNonExistent) Then
            '''' Expected error: FileNotFound
            this.ErrDescription = _
                "Databse file is not found in the specified folder." & _
                vbNewLine & "Source: " & FilePathName
            Err.Raise ErrNo.FileNotFoundErr, "LiteFSCheck", this.ErrDescription
        End If
        
        '''' Code below is executed if a new db is requested
        
        '''' Verify that a filename is provided
        If fso.FolderExists(FilePathName) Then
            '''' Expected error: FileNotFound
            this.ErrDescription = _
                "Filename is not provided or provided name conflicts with " & _
                "existing folder." & vbNewLine & "Source: " & FilePathName
            Err.Raise ErrNo.FileNotFoundErr, "LiteFSCheck", this.ErrDescription
        End If
        
        '''' Verify that a new file can be created
        Dim TextFile As Scripting.TextStream
        this.ErrSource = " LiteFSCheck"
        this.ErrDescription = vbNewLine & "Cannot create a new file."
        Set TextFile = fso.CreateTextFile(FilePathName)
        this.ErrSource = vbNullString
        this.ErrDescription = vbNullString
        TextFile.Close
        fso.DeleteFile FilePathName, True
        
        ExistsAccesibleValid = FilePathName
    End If
    On Error GoTo 0
    
    Exit Function
    
FILEPATH_ERROR:
    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    With this
        .ErrNumber = Err.Number
        .ErrSource = Err.Source & .ErrSource
        If InStr(Err.Description, vbNewLine & "Source: ") > 0 Then
            .ErrDescription = Err.Description
        Else
            .ErrDescription = Err.Description & .ErrDescription & vbNewLine & _
                              "Source: " & FilePathName
        End If
    End With
    ExistsAccesibleValid = vbNullString
End Function


Public Sub Raise()
    With this
        If .ErrNumber <> 0 Then
            Err.Raise .ErrNumber, .ErrSource, .ErrDescription
        End If
    End With
End Sub
