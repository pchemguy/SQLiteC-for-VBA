VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCConnDemoFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides fixtures for testing and as usage patterns"
'@Folder "SQLiteC For VBA.Connection"
'@ModuleDescription "Provides fixtures for testing and as usage patterns"
'@PredeclaredId
Option Explicit

Private Type TSQLiteCConnDemoFix
    dbm As SQLiteC
    DbConn As SQLiteCConnection
    DbStmt As SQLiteCStatement
    FixPath As String
End Type
Private this As TSQLiteCConnDemoFix


Public Function Create() As SQLiteCConnDemoFix
    Dim Instance As SQLiteCConnDemoFix
    Set Instance = New SQLiteCConnDemoFix
    Instance.Init
    Set Create = Instance
End Function

Friend Sub Init()
    Dim DllPath As String
    DllPath = "Library\SQLiteCforVBA\dll\" & ARCH
    '@Ignore IndexedDefaultMemberAccess
    Set this.dbm = SQLiteC(DllPath)
    this.FixPath = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & "Fixtures" & PATH_SEP
End Sub

Public Property Get dbm() As SQLiteC
    Set dbm = this.dbm
End Property

Private Property Get ConnDb(ByVal DbPathName As String) As SQLiteCConnection
    If Not this.DbConn Is Nothing Then
        Set ConnDb = this.DbConn
    Else
        Dim dbm As SQLiteC
        Set dbm = this.dbm
        Dim DbConn As SQLiteCConnection
        Set DbConn = dbm.DbConnInit(DbPathName)
        Set this.DbConn = DbConn
        Set ConnDb = DbConn
    End If
End Property

Public Property Get StmtDb(ByVal ConnDb As SQLiteCConnection) As SQLiteCStatement
    If Not this.DbStmt Is Nothing Then
        Set StmtDb = this.DbStmt
    Else
        Dim dbm As SQLiteC
        Set dbm = this.dbm
        Dim DbStmt As SQLiteCStatement
        Set DbStmt = dbm.DbStmtInit(ConnDb)
        Set this.DbStmt = DbStmt
        Set StmtDb = DbStmt
    End If
End Property

Public Property Get ConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set ConnDbRegular = ConnDb(DbPathName)
End Property

Public Property Get ConnDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set ConnDbMemory = ConnDb(DbPathName)
End Property

Public Property Get ConnDbTemp() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set ConnDbTemp = ConnDb(DbPathName)
End Property

Public Property Get ConnDbInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set ConnDbInvalidPath = ConnDb(DbPathName)
End Property

Public Property Get ConnDbNotDb() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.FixPath & "NotDb.db"
    Set ConnDbNotDb = ConnDb(DbPathName)
End Property

Public Property Get ConnDbLockedDb() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.FixPath & "ACLLocked" & PATH_SEP & "LockedDb.db"
    Set ConnDbLockedDb = ConnDb(DbPathName)
End Property
