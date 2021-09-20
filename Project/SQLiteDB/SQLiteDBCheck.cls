VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteDBCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides access and integrity check functionality"
'@Folder "SQLiteDB"
'@ModuleDescription "Provides access and integrity check functionality"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
'@IgnoreModule ValueRequired
Option Explicit

Private Type TSQLiteDBCheck
    MainDB As String '''' Absolute file pathname to the main database.
End Type
Private this As TSQLiteDBCheck


'@DefaultMember
Public Function Create(ByVal Database As String) As SQLiteDBCheck
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteDBCheck
    Set Instance = New SQLiteDBCheck
    Instance.Init Database
    Set Create = Instance
End Function


Friend Sub Init(ByVal Database As String)
    If Database = ":memory:" Then
        this.MainDB = Database
    Else
        this.MainDB = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    End If
End Sub


'''' File access tests and basic integrity validation via FS OS functions.
'@Description "Runs SQLite database access and basic integrity checks via OS."
Public Function CheckAccessAndBasicIntegrity() As Boolean
Attribute CheckAccessAndBasicIntegrity.VB_Description = "Runs SQLite database access and basic integrity checks via OS."
    Dim FilePathName As String
    FilePathName = this.MainDB
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
    '@Ignore VariableNotUsed: .Count property is accessed to test DbFolder access
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
    '@Ignore ConstantNotUsed
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
    
    CheckAccessAndBasicIntegrity = True
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
''''   >>> ?SQLiteDBCheck("Library\SQLiteDBVBA\SQLiteDBVBA.db").CheckIntegrity
''''   -- Integrity check passed for: '<Path>\SQLiteDBVBA.db'
''''   True
''''
''''   >>> ?SQLiteDBCheck("Library\SQLiteDBVBA\ICfailFKCfail.db").CheckIntegrity
''''   Error:
''''       message: Integrity check failed
''''
''''   >>> ?SQLiteDBCheck("Library\SQLiteDBVBA\ICokFKCfail.db").CheckIntegrity
''''   Error:
''''       message: Foreign key check failed
''''
'@Description "Runs SQLite database integrity checks via ADODB/SQL."
Public Function CheckIntegrity() As Boolean
Attribute CheckIntegrity.VB_Description = "Runs SQLite database integrity checks via ADODB/SQL."
    Const SQLITE_NOTADB_MESSAGE As String = "file is not a database (26)"
    
    Dim SQLQuery As String
    Dim CheckResult As Boolean
    Dim AdoRecordset As ADODB.Recordset
    
    Dim ADOTool As MinADO
    Set ADOTool = MinADO(this.MainDB)
    Dim SQLInfo As SQLiteSQLDbInfo
    '@Ignore DefaultMemberRequired: false positive
    Set SQLInfo = SQLiteSQLDbInfo()

    '''' 1) Verify that existing file is a database and only one db is atttached
    SQLQuery = SQLInfo.Databases
    Dim AttachedDbCount As Long
    Dim ErrNumber As Long
    Dim ErrSource As String
    Dim ErrDescription As String
    
    On Error Resume Next
    AttachedDbCount = ADOTool.GetAdoRecordset(SQLQuery).RecordCount
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    
    If ErrNumber <> 0 Then
        If (ErrNumber = ErrNo.AdoInvalidFileFormatErr) _
                And (ErrDescription = SQLITE_NOTADB_MESSAGE) Then
            Err.Raise ErrNo.AdoInvalidFileFormatErr, "SQLiteDB", _
                      "File is not a database"
        Else
            With Err
                .Raise ErrNumber, ErrSource, ErrDescription
            End With
        End If
    End If
    If AttachedDbCount <> 1 Then
        Err.Raise ErrNo.CustomErr, "SQLiteDB", _
                  "Please have exactly one database attached before the check."
    End If
    
    '''' 2) "PRAGMA integrity_check"
    SQLQuery = SQLInfo.CheckIntegrity
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.Fields("integrity_check") = "ok")
    Guard.ExpressionErr CheckResult, ErrNo.IntegrityCheckErr, _
                        "SQLiteDB", "Integrity check failed"
    
    '''' 3) "PRAGMA foreign_key_check"
    SQLQuery = SQLInfo.CheckFKs
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 0)
    Guard.ExpressionErr CheckResult, ErrNo.ConsistencyCheckErr, _
                        "SQLiteDB", "Foreign key check failed"

    Debug.Print "-- Integrity check passed for: '" & this.MainDB & "'"
    CheckIntegrity = True
End Function
