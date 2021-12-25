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
''''
'''' TODO: Extend CreateConnection/SQLiteCConnection:
''''         Define pass additional options to set appropriate OPEN_FLAGS
''''         AllowNonExistent = false - remove creatable flag
''''         Add OpenReadOnly, open in memory (?), cache control
''''         https://www.sqlite.org/uri.html
''''         https://www.sqlite.org/c3ref/open.html#urifilenamesinsqlite3open
''''
Option Explicit
Private Const MODULE_NAME As String = "SQLiteC"

#If VBA7 Then
'''' Init and shutdown, no db is necessary
Private Declare PtrSafe Function sqlite3_initialize Lib "SQLite3" () As Long
Private Declare PtrSafe Function sqlite3_shutdown Lib "SQLite3" () As Long
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Cache sharing, no db is necessary
Private Declare PtrSafe Function sqlite3_enable_shared_cache Lib "SQLite3" (ByVal Enable As Long) As Long
'''' Backup
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, _
    ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As LongPtr
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else
'''' Init and shutdown, no db is necessary
Private Declare Function sqlite3_initialize Lib "SQLite3" () As Long
Private Declare Function sqlite3_shutdown Lib "SQLite3" () As Long
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Cache sharing, no db is necessary
Private Declare Function sqlite3_enable_shared_cache Lib "SQLite3" (ByVal Enable As Long) As Long
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
    '''' The first created connection is designated as the main db, MainDB and
    '''' is set to this connection's ID, which is the pathname of its 'main' db.
    MainDB As Variant
    Functions As Scripting.Dictionary
    Modules As Scripting.Dictionary
    Pragmas As Scripting.Dictionary
    CompileOptions As Scripting.Dictionary
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
        #If Win64 Then
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
End Sub

'''' SQLiteC class hierarchy includes multiple instances of circular references,
'''' (parent object holding references to its children and children keeping a
'''' parent reference (see class diagram in the project docs). Such objects
'''' cannot be disposed of properly automatically by VBA.
''''
'''' SQLiteC is the top-level class and its encapsulated class SQLiteCConnection
'''' does not need and does not hold a reference to the manager. Thus, SQLiteC
'''' objects are destructed automatically, and through its Class_Terminate, this
'''' routine initiates a descending cascade of cleanup routines responsible for
'''' unraveling the circular references.
''''
Private Sub Class_Terminate()
    Logger.Logg "======== SQLiteC Class_Terminate ========", , DEBUGLEVEL_INFO
    With this
        If .Connections Is Nothing Then Exit Sub
        Dim DbConn As SQLiteCConnection
        Dim ConnName As Variant
        For Each ConnName In .Connections.Keys
            Set DbConn = .Connections(ConnName)
            DbConn.CleanUp
        Next ConnName
        Set DbConn = Nothing
        .Connections.RemoveAll
        Set .Connections = Nothing
    End With
    Dim ResultCode As SQLiteResultCodes
    ResultCode = Shutdown()
    If ResultCode <> SQLITE_OK Then Err.Raise ResultCode, _
        MODULE_NAME & "/Class_Terminate", "Shutdown error."
End Sub

Public Property Get MainDbId() As Variant
    MainDbId = this.MainDB
End Property

Public Property Get DllMan() As DllManager
    Set DllMan = this.DllMan
End Property

'@Description "Returns a dictionary with keys set to available SQLite functions."
Public Property Get Functions() As Scripting.Dictionary
    If this.Functions Is Nothing Then
        With LiteMetaADO(CreateILiteADO(":mem:"))
            Set this.Functions = .PragmaOptionsDict("function_list")
        End With
    End If
    Set Functions = this.Functions
End Property

'@Description "Returns a dictionary with keys set to available SQLite pragmas."
Public Property Get Pragmas() As Scripting.Dictionary
    If this.Pragmas Is Nothing Then
        With LiteMetaADO(CreateILiteADO(":mem:"))
            Set this.Pragmas = .PragmaOptionsDict("pragma_list")
        End With
    End If
    Set Pragmas = this.Pragmas
End Property

'@Description "Returns a dictionary with keys set to available SQLite modules."
Public Property Get Modules() As Scripting.Dictionary
    If this.Modules Is Nothing Then
        With LiteMetaADO(CreateILiteADO(":mem:"))
            Set this.Modules = .PragmaOptionsDict("module_list")
        End With
    End If
    Set Modules = this.Modules
End Property

'@Description "Returns a dictionary with keys set to available SQLite compile options."
Public Property Get CompileOptions() As Scripting.Dictionary
    If this.CompileOptions Is Nothing Then
        With LiteMetaADO(CreateILiteADO(":mem:"))
            Set this.CompileOptions = .PragmaOptionsDict("compile_options")
        End With
    End If
    Set CompileOptions = this.CompileOptions
End Property

'''' Reference: https://www.sqlite.org/c3ref/initialize.html
''''
'@Description "Initializes the library. No connection is necessary."
Public Function Initialize() As SQLiteResultCodes
    Initialize = sqlite3_initialize()
End Function

'''' Reference: https://www.sqlite.org/c3ref/initialize.html
''''
'@Description "Shuts down the library. No connection is necessary."
Public Function Shutdown() As SQLiteResultCodes
    Shutdown = sqlite3_shutdown()
End Function

'''' vbNullString is an acceptable DbPathName.
'''' Raises an error if DbPathName (or resolved value) has already been used.
'@Description "Creates a new SQLiteCConnection instance."
Public Function CreateConnection(ByVal DbPathName As String, _
                        Optional ByVal AllowNonExistent As Variant = True _
                        ) As SQLiteCConnection
Attribute CreateConnection.VB_Description = "Creates a new SQLiteCConnection instance."
    Dim PathCheck As LiteFSCheck
    Set PathCheck = LiteFSCheck(DbPathName, AllowNonExistent)
    Guard.ExpressionTrueErr PathCheck.ErrNumber <> 0, PathCheck.ErrNumber, _
        "SQLiteCConnection/Init", PathCheck.ErrDescription
    Dim FilePathName As String
    FilePathName = PathCheck.DatabasePathName
    If this.Connections.Exists(FilePathName) Then
        Err.Raise ErrNo.KeyAlreadyExistsErr, "SQLiteC", _
                  "Connection pathname already exists!"
    End If
    Dim DbConn As SQLiteCConnection
    Set DbConn = SQLiteCConnection(FilePathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.ObjectCreateErr, _
        "SQLiteC", "Failed to create an SQLiteCConnection instance."
    Set this.Connections(FilePathName) = DbConn
    If IsNull(this.MainDB) Then this.MainDB = FilePathName
    Set CreateConnection = DbConn
End Function

'@Description "Creates a new SQLiteCConnection instance."
Public Function CreateILiteADO(ByVal DbPathName As String, _
                      Optional ByVal AllowNonExistent As Variant = True _
                      ) As ILiteADO
Attribute CreateILiteADO.VB_Description = "Creates a new SQLiteCConnection instance."
    Dim dbc As SQLiteCConnection
    Set dbc = ConnDb(DbPathName)
    If dbc Is Nothing Then Set dbc = CreateConnection(DbPathName, AllowNonExistent)
    Set CreateILiteADO = dbc.ExecADO
End Function

'''' vbNullString is an acceptable DbPathName (should resolve to anon temp db).
'@Description "Returns an existing SQLiteCConnection instance or Nothing."
Public Function ConnDb(ByVal DbPathName As String) As SQLiteCConnection
Attribute ConnDb.VB_Description = "Returns an existing SQLiteCConnection instance or Nothing."
    If this.Connections.Exists(DbPathName) Then
        Set ConnDb = this.Connections(DbPathName)
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
'@Description "Returns SQLite version. No connection is necessary."
Public Function Version(Optional ByVal Numeric As Boolean = True) As Variant
Attribute Version.VB_Description = "Returns SQLite version. No database is necessary."
    If Numeric Then
        Version = sqlite3_libversion_number()
    Else
        Version = UTFlib.StrFromUTF8Ptr(sqlite3_libversion())
    End If
End Function

'''' Reference: https://www.sqlite.org/sharedcache.html
'''' Reference: https://www.sqlite.org/c3ref/enable_shared_cache.html
''''
'@Description "Enables and disables shared cache (process-wide)."
Public Function SetSharedCacheState(ByVal Enable As Boolean) As SQLiteResultCodes
    SetSharedCacheState = sqlite3_enable_shared_cache(IIf(Enable, 1, 0))
End Function


'#If VBA7 Then
'Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
'#Else
'Private Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
'Private Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
'#End If
