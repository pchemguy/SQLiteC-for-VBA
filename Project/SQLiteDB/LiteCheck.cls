VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteCheck"
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
Option Explicit

Private Type TLiteCheck
    MainDB As String '''' Absolute file pathname to the main database.
End Type
Private this As TLiteCheck


'''' Args:
''''   Database (string):
''''     Absolute path name to the database. No path processing should be done
''''       here. E.g., relative path will result in the PathNotFound error.
'@DefaultMember
Public Function Create(ByVal Database As String) As LiteCheck
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteCheck
    Set Instance = New LiteCheck
    Instance.MainDB = Database
    Set Create = Instance
End Function


Public Property Get MainDB() As String
    MainDB = this.MainDB
End Property


Public Property Let MainDB(ByVal Database As String)
    this.MainDB = Database
End Property


'@Description "Checks that provided path exists and is accessible."
Public Function PathExistsAccessible(ByVal FilePathName As String) As Boolean
Attribute PathExistsAccessible.VB_Description = "Checks that provided path exists and is accessible."
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
Attribute FileAccessibleValid.VB_Description = "Verifies that the database file is accessible and valid."
    Dim fso As New Scripting.FileSystemObject
    
    '''' Verify that the file size is >=100 (SQLite header size is 100 bytes).
    Dim DbFile As Scripting.File
    Set DbFile = fso.GetFile(FilePathName)
    If DbFile.Size < 100 Then
        Err.Raise ErrNo.AdoInvalidFileFormatErr, "LiteCheck", "File is not " & _
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
        Err.Raise ErrNo.AdoInvalidFileFormatErr, "LiteCheck", "Database " & _
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
Public Function ExistsAccesibleValid() As Boolean
Attribute ExistsAccesibleValid.VB_Description = "Runs SQLite database access and basic integrity checks via OS."
    If this.MainDB = ":memory:" Then
        ExistsAccesibleValid = True
        Exit Function
    End If
    
    Dim fso As New Scripting.FileSystemObject
    If fso.FileExists(this.MainDB) Then
        ExistsAccesibleValid = FileAccessibleValid(this.MainDB)
    Else
        ExistsAccesibleValid = PathExistsAccessible(this.MainDB)
    End If
End Function


'''' Integrity tests executed by the SQLite engine initiated via ADODB/SQL.
''''
'''' The database to be checked should be attached as the "main" and only
'''' database. While integrity/FK check should run against the specified schema,
'''' if provided, I had issues with getting expected behavior. Prefer to run
'''' checks when a single database being checked is attached as "main".
'''' 1) Verify that existing file is a database and exactly one db is attached.
'''' 2) Run integrity check and verify that it returns "ok"
'''' 3) Run foreign key check and verify that nothing is returned.
''''
'''' Examples:
''''   >>> ?liteCheck("Library\SQLiteDBVBA\SQLiteDBVBA.db").IntegrityADODB
''''   -- Integrity check passed for: '<Path>\SQLiteDBVBA.db'
''''   True
''''
''''   >>> ?liteCheck("Library\SQLiteDBVBA\ICfailFKCfail.db").IntegrityADODB
''''   Error:
''''       message: Integrity check failed
''''
''''   >>> ?liteCheck("Library\SQLiteDBVBA\ICokFKCfail.db").IntegrityADODB
''''   Error:
''''       message: Foreign key check failed
''''
'@Description "Runs SQLite database integrity checks via ADODB/SQL."
Public Function IntegrityADODB() As Boolean
Attribute IntegrityADODB.VB_Description = "Runs SQLite database integrity checks via ADODB/SQL."
    Dim ADOTool As ILiteADO
    Set ADOTool = LiteADO(this.MainDB)
    Dim SQLInfo As SQLiteSQLDbInfo
    '@Ignore DefaultMemberRequired: false positive
    Set SQLInfo = SQLiteSQLDbInfo()

    '''' 1) Verify that existing file is a database and only one db is atttached
    '''' Expected error: AdoInvalidFileFormatErr due to damaged database file.
    '''' Database file damage may not be detected until the first query.
    On Error GoTo INTEGRITY_ERROR
    Dim AttachedDbCount As Long
    AttachedDbCount = ADOTool.GetAdoRecordset(SQLInfo.Databases).RecordCount
    On Error GoTo 0
    
    If AttachedDbCount <> 1 Then
        Err.Raise ErrNo.CustomErr, "LiteCheck", _
                  "Please have exactly one database attached before the check."
    End If
    
    '''' 2) "PRAGMA integrity_check"
    On Error GoTo INTEGRITY_ERROR
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLInfo.CheckIntegrity)
    On Error GoTo 0
    If Not AdoRecordset.Fields("integrity_check") = "ok" Then
        Err.Raise ErrNo.IntegrityCheckErr, "LiteCheck", "Integrity check failed"
    End If
    
    '''' 3) "PRAGMA foreign_key_check"
    On Error GoTo INTEGRITY_ERROR
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLInfo.CheckFKs)
    On Error GoTo 0
    If Not AdoRecordset.RecordCount = 0 Then
        Err.Raise ErrNo.ConsistencyCheckErr, "LiteCheck", "Foreign key check failed"
    End If

    Debug.Print "-- Integrity check passed for: '" & this.MainDB & "'"
    IntegrityADODB = True
    Exit Function
    
INTEGRITY_ERROR:
    Const SQLITE_NOTADB_MESSAGE As String = "file is not a database (26)"
    If (Err.Number = ErrNo.AdoInvalidFileFormatErr) _
            And (Err.Description = SQLITE_NOTADB_MESSAGE) Then
        Err.Raise ErrNo.AdoInvalidFileFormatErr, "LiteCheck", _
                  "File is not a database"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & _
                  "Unexpected IntegrityADODB error."
    End If
End Function
