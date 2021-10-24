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
'@Folder "SQLite.SQLiteC For VBA.Manager"
'@ModuleDescription "Provides common workflows for SQLite db interactions"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Backup
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, _
    ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Backup
Private Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, _
    ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
#End If

Option Explicit

Private Type TSQLiteC
    DllMan As DllManager
    DbConnMan As Scripting.Dictionary
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
    Set this.DbConnMan = New Scripting.Dictionary
    this.DbConnMan.CompareMode = TextCompare
    this.MainDB = Null
End Sub


Public Property Get MainDbId() As Variant
    MainDbId = this.MainDB
End Property


Public Property Get DllMan() As DllManager
    Set DllMan = this.DllMan
End Property


'''' vbNullString is an acceptable StmtName.
'''' Raises an error if DbPathName (or resolved value) has already been used.
'@Description "Creates a new SQLiteCConnection instance."
Public Function CreateConnection(ByVal DbPathName As String) As SQLiteCConnection
Attribute CreateConnection.VB_Description = "Creates a new SQLiteCConnection instance."
    If this.DbConnMan.Exists(DbPathName) Then Err.Raise KeyAlreadyExistsErr, _
        "SQLiteC", "Connection pathname already exists!"
    Dim DbConn As SQLiteCConnection
    Set DbConn = SQLiteCConnection(DbPathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteC", "Failed to create an SQLiteCConnection instance."
    Dim DbConnId As String
    DbConnId = DbConn.DbPathName
    If this.DbConnMan.Exists(DbConnId) Then Err.Raise KeyAlreadyExistsErr, _
        "SQLiteC", "Connection pathname already exists!"
    If IsNull(this.MainDB) Then this.MainDB = DbConnId
    Set this.DbConnMan(DbConnId) = DbConn
    Set CreateConnection = DbConn
End Function


'''' vbNullString is an acceptable DbPathName (should resolve to anon temp db).
'@Description "Returns an existing SQLiteCConnection instance or Nothing."
Public Function ConnDb(Optional ByVal DbPathName As String = vbNullString _
                      ) As SQLiteCConnection
Attribute ConnDb.VB_Description = "Returns an existing SQLiteCConnection instance or Nothing."
    '''' SQLiteCConnection is responsible for DbPathName validation/resolution
    If this.DbConnMan.Exists(DbPathName) Then
        Set ConnDb = this.DbConnMan(DbPathName)
    ElseIf Len(DbPathName) = 0 And this.DbConnMan.Exists(this.MainDB) Then
        Set ConnDb = this.DbConnMan(this.MainDB)
    Else
        Set ConnDb = Nothing
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/backup_finish.html
'''' Reference: https://www.sqlite.org/backup.html
'''' Reference: https://www.sqlite.org/lang_vacuum.html#vacuuminto
''''
'@Description "Performs online database backup."
Public Function DupDbOnlineFull(ByVal DbConnDest As SQLiteCConnection, _
                       Optional ByVal DstAlias As String = "main", _
                       Optional ByVal DbConnSrc As SQLiteCConnection, _
                       Optional ByVal SrcAlias As String = "main") As SQLiteResultCodes
Attribute DupDbOnlineFull.VB_Description = "Performs online database backup."
    Guard.NullReference DbConnSrc
    Dim DbConnDst As SQLiteCConnection
    Set DbConnDst = IIf(DbConnDest Is Nothing, DbConnSrc, DbConnDest)
    #If VBA7 Then
        Dim DbHandleBak As LongPtr
    #Else
        Dim DbHandleBak As Long
    #End If
    DbHandleBak = sqlite3_backup_init(DbConnDst.DbHandle, DstAlias, _
                                      DbConnSrc.DbHandle, SrcAlias)
    Dim ResultCode As SQLiteResultCodes
    If DbHandleBak = 0 Then GoTo RESULT_CODE:
    ResultCode = sqlite3_backup_step(DbHandleBak, -1)
    If ResultCode <> SQLITE_DONE Then GoTo RESULT_CODE:
    ResultCode = sqlite3_backup_finish(DbHandleBak)
    Debug.Assert ResultCode = SQLITE_OK
    
RESULT_CODE:
    ResultCode = DbConnDst.ErrInfoRetrieve
    DupDbOnlineFull = ResultCode
    Exit Function
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
