VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteACID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "This module provides ACID related functionality"
'@Folder "SQLiteDB"
'@ModuleDescription "This module provides ACID related functionality"
'@PredeclaredId
Option Explicit

Private Const SQLITE_NOTADB_MESSAGE As String = "file is not a database (26)"
Private Const SQLITE_IOERR_MESSAGE As String = "disk I/O error (10)"

Private Type TLiteACID
    ADOTool As ILiteADO
    SQLInfo As LiteMetaSQL
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
    Set this.SQLInfo = LiteMetaSQL()
End Sub


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
'@Description "Runs SQLite database integrity checks via ADODB/SQL."
Friend Function IntegrityADODB() As Boolean
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

    Debug.Print "-- Integrity check passed for: '" & ADOTool.MainDB & "'"
    IntegrityADODB = True
    Exit Function
    
INTEGRITY_ERROR:
    If (Err.Number = ErrNo.OLE_DB_ODBC_Err) _
            And (Err.Description = SQLITE_NOTADB_MESSAGE) Then
        Err.Raise ErrNo.OLE_DB_ODBC_Err, "LiteCheck", _
                  "File is not a database"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description & vbNewLine & _
                  "Unexpected IntegrityADODB error."
    End If
End Function
