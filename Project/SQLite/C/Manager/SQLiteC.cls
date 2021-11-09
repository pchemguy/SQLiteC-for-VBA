VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides common workflows for SQLite db interactions"
'@Folder "SQLite.C.Manager"
'@ModuleDescription "Provides common workflows for SQLite db interactions"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess

Option Explicit

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Backup
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, _
    ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Backup
Private Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, _
    ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
#End If

Private Type TSQLiteC
    DllMan As DllManager
    Connections As Scripting.Dictionary
    SkipCleanUp As Boolean
    '''' The first created connection is designated as the main db, MainDB and
    '''' is set to this connection's ID, which is the pathname of its 'main' db.
    MainDB As Variant
End Type
Private this As TSQLiteC


'@DefaultMember
Public Function Create(ByVal DllPath As String, _
              Optional ByVal DllNames As Variant = Empty) As SQLiteC
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteC
    Set Instance = New SQLiteC
    Instance.Init DllPath, DllNames
    Set Create = Instance
End Function

Friend Sub Init(ByVal DllPath As String, _
       Optional ByVal DllNames As Variant = Empty)
    Dim FileNames As Variant
    If Not IsEmpty(DllNames) Then
        FileNames = DllNames
    Else
        #If WIN64 Then
            '''' SQLite3.dll-x64 is built with MSYS2/MinGWx64
            '''' LoadLibrary resolves/loads dependencies automatically.
            FileNames = "sqlite3.dll"
        #Else
            '''' SQLite3.dll-x32 is built with MSVC and follows STDCALL
            '''' LoadLibrary fails to resolve/load dependencies automatically,
            '''' so loading them explicitly.
            FileNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                              "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
    End If
    '''' DllManager is responsible for DllPath validation/resolution
    Set this.DllMan = DllManager.Create(DllPath, FileNames)
    Set this.Connections = New Scripting.Dictionary
    this.Connections.CompareMode = TextCompare
    this.MainDB = Null
    this.SkipCleanUp = False
End Sub

Private Sub Class_Terminate()
    Dim LogMsg As String
    LogMsg = "========== SQLiteC Class_Terminate ==========" & vbNewLine & _
             "Cleanup " & IIf(SkipCleanUp, "skipped", "performed")
    Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    If Not this.SkipCleanUp Then CleanUp
End Sub

'''' SQLiteC class hierarchy includes multiple instances of circular references,
'''' (parent object holding references to its children and children keeping a
'''' parent reference (see class diagram in the project docs). Such objects
'''' cannot be disposed of properly automatically by VBA.
''''
'''' SQLiteC is the top-level class and its encapsulated class SQLiteCConnection
'''' does not need and does not hold a reference to the manager. Thus, SQLiteC
'''' objects are destructed automatically, and through its Class_Terminate, this
'''' routines initiates a descending cascade of cleanup routines responsible for
'''' unraveling the circular references.
''''
'''' FixObjC is extensively used as a fixture for testing the SQLiteC subpackage.
'''' For testing purposes, a single manager is normally sufficient. SQLiteC is
'''' responsible for loading DLL libraries, so its instantiation is somewhat
'''' expensive. FixObjC class instantiates SQLiteC class the first time it is
'''' requested and reuses the manager. To minimize undesired test interactions,
'''' however, the Cleanup routine is called to clear all descendant objects, and
'''' most of the time it works fine. In case of SQLiteC.DupDbOnlineFull, however,
'''' the current test use two independent SQLiteCConnection objects generated via
'''' FixObjC. That means the when the second object is requested, the first one
'''' is by the FixObjC.GetDBM initiated cleanup cascade. To permit an exception,
'''' the SkipCleanUp attribute is introduced into the SQLiteC class. If set to
'''' true SQLiteC.Cleanup cascade is aborted. This attribute, in turn is exposed
'''' on dbm, and FixObjC exposes its dbm object, making it possible to perform
'''' the test (see SQLiteCTests.ztcDupDbOnlineFull_*).
Friend Sub CleanUp()
    If SkipCleanUp Or this.Connections Is Nothing Then Exit Sub
    Dim DbConn As SQLiteCConnection
    Dim ConnName As Variant
    For Each ConnName In this.Connections.Keys
        Set DbConn = this.Connections(ConnName)
        DbConn.CleanUp
        this.Connections.Remove ConnName
    Next ConnName
End Sub

Friend Property Get SkipCleanUp() As Boolean
    SkipCleanUp = this.SkipCleanUp
End Property

'@Ignore ProcedureNotUsed
Friend Property Let SkipCleanUp(ByVal Value As Boolean)
    this.SkipCleanUp = Value
End Property

Public Property Get MainDbId() As Variant
    MainDbId = this.MainDB
End Property

Public Property Get DllMan() As DllManager
    Set DllMan = this.DllMan
End Property

'''' vbNullString is an acceptable StmtName.
'''' Raises an error if DbPathName (or resolved value) has already been used.
'@Description "Creates a new SQLiteCConnection instance."
Public Function CreateConnection(ByVal DbPathName As String, _
                        Optional ByVal AllowNonExistent As Variant = True _
                        ) As SQLiteCConnection
Attribute CreateConnection.VB_Description = "Creates a new SQLiteCConnection instance."
    If this.Connections.Exists(DbPathName) Then Err.Raise ErrNo.KeyAlreadyExistsErr, _
        "SQLiteC", "Connection pathname already exists!"
    Dim DbConn As SQLiteCConnection
    Set DbConn = SQLiteCConnection(DbPathName, AllowNonExistent)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteC", "Failed to create an SQLiteCConnection instance."
    Dim DbConnId As String
    DbConnId = DbConn.DbPathName
    If this.Connections.Exists(DbConnId) Then Err.Raise ErrNo.KeyAlreadyExistsErr, _
        "SQLiteC", "Connection pathname already exists!"
    If IsNull(this.MainDB) Then this.MainDB = DbConnId
    Set this.Connections(DbConnId) = DbConn
    Set CreateConnection = DbConn
End Function

'''' vbNullString is an acceptable DbPathName (should resolve to anon temp db).
'@Description "Returns an existing SQLiteCConnection instance or Nothing."
Public Function ConnDb(Optional ByVal DbPathName As String = vbNullString _
                      ) As SQLiteCConnection
Attribute ConnDb.VB_Description = "Returns an existing SQLiteCConnection instance or Nothing."
    '''' SQLiteCConnection is responsible for DbPathName validation/resolution
    If this.Connections.Exists(DbPathName) Then
        Set ConnDb = this.Connections(DbPathName)
    ElseIf Len(DbPathName) = 0 And this.Connections.Exists(this.MainDB) Then
        Set ConnDb = this.Connections(this.MainDB)
    Else
        Set ConnDb = Nothing
    End If
End Function

'''' Reference: https://www.sqlite.org/c3ref/backup_finish.html
'''' Reference: https://www.sqlite.org/backup.html
'''' Reference: https://www.sqlite.org/lang_vacuum.html#vacuuminto
''''
'''' Returns:
''''    number of pages copied
''''
'@Description "Performs online database backup."
Public Function DupDbOnlineFull(ByVal DbConnDest As SQLiteCConnection, _
                       Optional ByVal DstAlias As String = "main", _
                       Optional ByVal DbConnSrc As SQLiteCConnection, _
                       Optional ByVal SrcAlias As String = "main") As Long
Attribute DupDbOnlineFull.VB_Description = "Performs online database backup."
    Dim DbConnDst As SQLiteCConnection
    Set DbConnDst = IIf(DbConnDest Is Nothing, DbConnSrc, DbConnDest)
    FixGuard.DbNotOpened DbConnSrc, "SQLiteC/DupDbOnlineFull"
    FixGuard.DbNotOpened DbConnDst, "SQLiteC/DupDbOnlineFull"
    #If VBA7 Then
        Dim DbHandleBak As LongPtr
        Dim SrcAliasPtr As LongPtr
        Dim DstAliasPtr As LongPtr
    #Else
        Dim DbHandleBak As Long
        Dim SrcAliasPtr As Long
        Dim DstAliasPtr As Long
    #End If
    
    Dim SrcAliasUTF8B() As Byte
    Dim DstAliasUTF8B() As Byte
    SrcAliasUTF8B = UTFlib.UTF8BytesFromStr(SrcAlias)
    DstAliasUTF8B = UTFlib.UTF8BytesFromStr(DstAlias)
    SrcAliasPtr = VarPtr(SrcAliasUTF8B(0))
    DstAliasPtr = VarPtr(DstAliasUTF8B(0))
    
    DbHandleBak = sqlite3_backup_init(DbConnDst.DbHandle, DstAliasPtr, _
                                      DbConnSrc.DbHandle, SrcAliasPtr)
    If DbHandleBak = 0 Then GoTo RESULT_CODE:
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_backup_step(DbHandleBak, -1)
    If ResultCode <> SQLITE_DONE Then GoTo RESULT_CODE:
    Dim PagesLeft As Long
    PagesLeft = sqlite3_backup_remaining(DbHandleBak)
    If PagesLeft <> 0 Then GoTo RESULT_CODE:
    Dim PagesDone As Long
    PagesDone = sqlite3_backup_pagecount(DbHandleBak)
    ResultCode = sqlite3_backup_finish(DbHandleBak)
    Debug.Assert ResultCode = SQLITE_OK
    DupDbOnlineFull = PagesDone
    Exit Function
RESULT_CODE:
    '@Ignore AssignmentNotUsed
    ResultCode = DbConnDst.ErrInfoRetrieve
    DupDbOnlineFull = 0
End Function

'''' Reference: https://www.sqlite.org/c3ref/libversion.html
''''
'@Description "Returns SQLite version. No database is necessary."
Public Function Version(Optional ByVal Numeric As Boolean = True) As Variant
Attribute Version.VB_Description = "Returns SQLite version. No database is necessary."
    If Numeric Then
        Version = sqlite3_libversion_number()
    Else
        Version = UTFlib.StrFromUTF8Ptr(sqlite3_libversion())
    End If
End Function


'#If VBA7 Then
'Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'#Else
'Private Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
'Private Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
'#End If
