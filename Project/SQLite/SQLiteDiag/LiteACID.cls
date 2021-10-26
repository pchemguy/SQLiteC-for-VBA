VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteACID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "This module provides ACID related functionality"
'@Folder "SQLite.SQLiteDiag"
'@ModuleDescription "This module provides ACID related functionality"
'@PredeclaredId
'@Exposed
'@IgnoreModule ValueRequired: false positive with ADODB
'@IgnoreModule ProcedureNotUsed
''''
'''' This module provides extended integrity checks via the SQLite engine
'''' to be performed using a dedicated ADODB Connection object with the target
'''' database being the attached as the "main" and only database. (While those
'''' checks should should run on specific alias, I could not get it to work
'''' properly.) All these tests can be run on a transaction locked database,
'''' and they should not introduce any locks.
''''
'''' The main entry is IntegrityADODB. SQLite provides two pragmas for
'''' consistency checks, one performs a series of checks, except for the
'''' foreign key constraints checks, and the other one covers foreing keys.
''''
'''' Transaction-related lock checks should be performed before starting a
'''' a transaction or attempting a modification with auto transaction. These
'''' should use the primary Connection object and for this reason should not
'''' be a part of this module.
''''
Option Explicit

Private Const JOURNAL_MODES As String = "| DELETE | TRUNCATE | WAL |"
Private Const SQLITE_READONLY_MSG As String = "attempt to write a readonly database (8)"
Private Const SQLITE_READONLY_CODE As Long = 8
Private Const SQLITE_IOERR_MSG As String = "disk I/O error (10)"
Private Const SQLITE_IOERR_CODE As Long = 10
Private Const SQLITE_NOTADB_MSG As String = "file is not a database (26)"
Private Const SQLITE_NOTADB_CODE As Long = 26

Private Type TLiteACID
    ADOTool As ILiteADO
    SQLInfo As LiteMetaSQL
    ErrNumber As Long
    ErrSource As String
    ErrDescription As String
    ErrStack As String
    ErrSQLiteNumber As Long
End Type
Private this As TLiteACID


'@DefaultMember
Public Function Create(ByVal ADOTool As ILiteADO) As LiteACID
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteACID
    Set Instance = New LiteACID
    Instance.Init ADOTool
    Set Create = Instance
End Function


Friend Sub Init(ByVal ADOTool As ILiteADO)
    Guard.NullReference ADOTool
    Guard.EmptyString ADOTool.MainDB
    
    Set this.ADOTool = ADOTool
    Set this.SQLInfo = LiteMetaSQL.Create()
End Sub


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

Public Property Get ErrSQLiteNumber() As Long
    ErrSQLiteNumber = this.ErrSQLiteNumber
End Property

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
''''   >>> ?LiteACID(LiteADO("Library\SQLiteDBVBA\SQLiteDBVBA.db")).IntegrityADODB
''''   -- Integrity check passed for: '<Path>\SQLiteDBVBA.db'
''''   True
''''
''''   >>> ?LiteACID(LiteADO("Library\SQLiteDBVBA\ICfailFKCfail.db")).IntegrityADODB
''''   Error:
''''       message: Integrity check failed
''''
''''   >>> ?LiteACID(LiteADO("Library\SQLiteDBVBA\ICokFKCfail.db")).IntegrityADODB
''''   Error:
''''       message: Foreign key check failed
''''
'''' https://www.sqlite.org/pragma.html#pragma_integrity_check
'''' https://www.sqlite.org/pragma.html#pragma_quick_check
'''' https://www.sqlite.org/pragma.html#pragma_foreign_key_check
''''
'@Description "Runs SQLite database integrity checks via ADODB/SQL."
Friend Function IntegrityADODB() As Boolean
Attribute IntegrityADODB.VB_Description = "Runs SQLite database integrity checks via ADODB/SQL."
    Const ThisProcName As String = "IntegrityADODB"
    
    Dim ADOTool As ILiteADO
    Set ADOTool = this.ADOTool
    Dim SQLInfo As LiteMetaSQL
    Set SQLInfo = this.SQLInfo

    '''' 1) Verify that existing file is a database and only one db is atttached
    '''' Expected error: OLE_DB_ODBC_Err due to damaged database file.
    '''' Database file damage may not be detected until the first query.
    On Error GoTo INTEGRITY_ERROR
    Dim AttachedDbCount As Long
    AttachedDbCount = ADOTool.GetAdoRecordset(SQLInfo.Databases).RecordCount
    On Error GoTo 0
    
    If AttachedDbCount <> 1 Then
        Err.Raise ErrNo.CustomErr, "LiteACID", _
                  "Please have exactly one database attached before the check."
    End If
    
    '''' 2) "PRAGMA integrity_check"
    On Error GoTo INTEGRITY_ERROR
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLInfo.CheckIntegrity)
    On Error GoTo 0
    '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
    If Not AdoRecordset.Fields("integrity_check") = "ok" Then
        Err.Raise ErrNo.IntegrityCheckErr, "LiteACID", "Integrity check failed"
    End If
    
    '''' 3) "PRAGMA foreign_key_check"
    On Error GoTo INTEGRITY_ERROR
    Set AdoRecordset = ADOTool.GetAdoRecordset(SQLInfo.CheckFKs)
    On Error GoTo 0
    If Not AdoRecordset.RecordCount = 0 Then
        Err.Raise ErrNo.ConsistencyCheckErr, "LiteACID", "Foreign key check failed"
    End If

    Debug.Print "-- Integrity check passed for: '" & ADOTool.MainDB & "'"
    IntegrityADODB = True
    Exit Function
    
INTEGRITY_ERROR:
    If Err.Number = ErrNo.OLE_DB_ODBC_Err Then
        Select Case Err.Description
            Case SQLITE_NOTADB_MSG
                this.ErrSQLiteNumber = SQLITE_NOTADB_CODE
                this.ErrDescription = Err.Description & vbNewLine & _
                                      "File is damaged or not a database."
            Case Else
                this.ErrSQLiteNumber = 1
        End Select
    Else
        this.ErrDescription = Err.Description & vbNewLine & "Unexpected error."
    End If
    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    Err.Raise Err.Number, "LiteACID", this.ErrDescription
End Function


Public Property Get JournalMode() As String
    JournalMode = UCase$(CStr(this.ADOTool.GetScalar("PRAGMA journal_mode")))
End Property


Public Property Let JournalMode(ByVal Mode As String)
    Guard.ExpressionErr InStr(JOURNAL_MODES, "| " & UCase$(Mode) & " |") > 0, _
                        ErrNo.InvalidParameterErr, "LiteACID", _
                        "Please use DELETE, TRUNCATE, or WAL. Received: " & Mode
    Const ThisProcName As String = "JournalModeLet"
    
    Dim dbm As ILiteADO
    Set dbm = this.ADOTool
    
    Dim SQLQuery As String
    SQLQuery = "PRAGMA journal_mode(" & Mode & ")"
    On Error GoTo WRITE_FAILED
        '@Ignore FunctionReturnValueDiscarded
        dbm.ExecuteNonQuery SQLQuery
    On Error GoTo 0
    
    Exit Property
    
WRITE_FAILED:
    If Err.Number = ErrNo.OLE_DB_ODBC_Err Then
        Select Case Err.Description
            Case SQLITE_READONLY_MSG
                this.ErrSQLiteNumber = SQLITE_READONLY_CODE
                this.ErrDescription = Err.Description & vbNewLine & _
                                      "Database is write locked."
            Case SQLITE_IOERR_MSG
                this.ErrSQLiteNumber = SQLITE_IOERR_CODE
                this.ErrDescription = Err.Description & vbNewLine & _
                                      "Database is write locked."
            Case Else
                this.ErrSQLiteNumber = 1
        End Select
    Else
        this.ErrDescription = Err.Description & vbNewLine & "Unexpected error."
    End If
    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    Err.Raise Err.Number, "LiteACID", this.ErrDescription
End Property


Public Function JournalModeToggle() As String
    Const ThisProcName As String = "JournalModeToggle"
    
    Dim JMode As String
    JMode = JournalMode
    If JMode = "WAL" Then
        JMode = "DELETE"
    Else
        JMode = "WAL"
    End If
    On Error GoTo TOGGLE_FAILED
        JournalMode = JMode
    On Error GoTo 0
    
    If JMode <> JournalMode Then
        Err.Raise ErrNo.AdoProviderFailedErr, "LiteACID", _
                  "Failed to change journal mode"
    End If
    
    JournalModeToggle = JMode
    Exit Function
    
TOGGLE_FAILED:
    Select Case Err.Number
        Case ErrNo.OLE_DB_ODBC_Err
            Select Case this.ErrSQLiteNumber
                Case SQLITE_READONLY_CODE, SQLITE_IOERR_CODE
                    this.ErrDescription = Err.Description
                Case Else
                    this.ErrSQLiteNumber = 1
            End Select
        Case ErrNo.AdoProviderFailedErr
            this.ErrDescription = Err.Description
        Case Else
            this.ErrDescription = Err.Description & vbNewLine & "Unexpected error."
    End Select
    
    this.ErrStack = ThisProcName & vbNewLine & this.ErrStack
    Err.Raise Err.Number, "LiteACID", this.ErrDescription
End Function


Public Function LockedInTransaction() As Boolean
    Dim JMode As String
    JMode = JournalMode
    
    Guard.ExpressionErr InStr(JOURNAL_MODES, "| " & UCase$(JMode) & " |") > 0, _
                        ErrNo.IncompatibleStatusErr, "LiteACID", _
                        "Please set DELETE, TRUNCATE, or WAL. Current: " & JMode
                        
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    Dim FilePathName As String
    Select Case JMode
        Case "WAL"
            FilePathName = this.ADOTool.MainDB & "-shm"
            If fso.FileExists(FilePathName) And fso.GetFile(FilePathName).Size > 0 Then
                Dim TextFile As Scripting.TextStream
                On Error Resume Next
                    Set TextFile = fso.OpenTextFile(FilePathName, ForReading)
                    TextFile.Read 1
                    If Err.Number <> 0 Then LockedInTransaction = True
                    If Not TextFile Is Nothing Then TextFile.Close
                On Error GoTo 0
            End If
        Case "DELETE", "TRUNCATE"
            FilePathName = this.ADOTool.MainDB & "-journal"
            If fso.FileExists(FilePathName) Then
                If fso.GetFile(FilePathName).Size > 0 Then LockedInTransaction = True
            End If
    End Select
End Function


Public Function LockedReadOnly() As Boolean
    If LockedInTransaction Then
        LockedReadOnly = True
        Exit Function
    End If
    
    On Error Resume Next
        '@Ignore FunctionReturnValueDiscarded
        JournalModeToggle
        If Err.Number <> 0 Then
            LockedReadOnly = True
        Else
            LockedReadOnly = False
            '@Ignore FunctionReturnValueDiscarded
            JournalModeToggle
        End If
    On Error GoTo 0
End Function

'    Dim JournalModeUC As String
'    JournalModeUC = UCase$(JournalMode)
'    Dim CheckResult As Boolean
'    CheckResult = CBool(InStr("| " & JOURNAL_MODES & " |", "| " & JournalModeUC & " |"))
'    Guard.Expression CheckResult, "SQLiteDB/JournalModeSet", _
'        "Invalid journal mode: '" & JournalMode & "'. Valid values are: " _
'        & JOURNAL_MODES
'
'    Dim SQLQuery As String
'    Dim Databases As Variant
'    If UCase$(SchemaAlias) <> "ALL" Then
'        Databases = Array(SchemaAlias)
'    Else
'        SQLQuery = "SELECT name FROM pragma_database_list"
'        Databases = ArrayLib.GetRow(GetAdoRecordset(SQLQuery).GetRows, 0)
'    End If
'

'Public Function DbIsWriteLocked() As Boolean
'    Dim dbm As ILiteADO
'    Set dbm = this.ADOTool
'
'    Dim JournalMode As String
'    JournalMode = dbm.GetScalar
'End Function



'Private Sub ztcExistsAccesibleValid_ThrowsOnBadMagicA()
'    Dim FilePathName As String
'    FilePathName = zfxFixturePrefix & "TestCWAL.db"
'    Dim dbm As ILiteADO
'    Set dbm = LiteADO(FilePathName)
'
'    Dim AdoConnection As ADODB.Connection
'    Set AdoConnection = dbm.AdoConnection
'
'    LiteCheck(FilePathName).ExistsAccesibleValid
'    Dim Response As Variant
'    Response = dbm.GetScalar("PRAGMA journal_mode")
'    dbm.ExecuteNonQuery "PRAGMA journal_mode='DELETE'"
'    Response = dbm.GetScalar("PRAGMA journal_mode")
'
'    On Error Resume Next
'    FilePathName = zfxFixturePrefix & "TestCWAL.db"
'    Set dbm = LiteADO(FilePathName)
'    dbm.ExecuteNonQuery "BEGIN IMMEDIATE"
'    LiteCheck(FilePathName & "-shm").ExistsAccesibleValid
'    dbm.ExecuteNonQuery "ROLLBACK"
'    Guard.AssertExpectedError Assert, ErrNo.TextStreamReadErr
'End Sub
