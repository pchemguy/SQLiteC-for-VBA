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
'@Folder "SQLiteC For VBA.Manager"
'@ModuleDescription "Provides common workflows for SQLite db interactions"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As LongPtr, _
    ByVal zDestName As LongPtr, ByVal hDbSource As LongPtr, ByVal zSourceName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As LongPtr, ByVal nPage As Long) As Long
Private Declare PtrSafe Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As LongPtr) As Long
#Else
Private Declare Function sqlite3_backup_init Lib "SQLite3" (ByVal hDbDest As Long, _
    ByVal zDestName As Long, ByVal hDbSource As Long, ByVal zSourceName As Long) As Long
Private Declare Function sqlite3_backup_step Lib "SQLite3" (ByVal hBackup As Long, ByVal nPage As Long) As Long
Private Declare Function sqlite3_backup_finish Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_remaining Lib "SQLite3" (ByVal hBackup As Long) As Long
Private Declare Function sqlite3_backup_pagecount Lib "SQLite3" (ByVal hBackup As Long) As Long
#End If

Option Explicit

Private Type TSQLiteC
    DllMan As DllManager
    DbConn As SQLiteCConnection
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
            FileNames = "sqlite3.dll"
        #Else
            FileNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                              "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
    End If
    '''' DllManager is responsible for DllPath validation/resolution
    Set this.DllMan = DllManager.Create(DllPath, FileNames)
End Sub


Public Property Get DllMan() As DllManager
    Set DllMan = this.DllMan
End Property


Public Property Get DbConnGet() As SQLiteCConnection
    Set DbConnGet = this.DbConn
End Property


Public Function DbConnInit(ByVal DbPathName As String) As SQLiteCConnection
    '''' SQLiteCConnection is responsible for DbPathName validation/resolution
    Dim DbConn As SQLiteCConnection
    Set DbConn = SQLiteCConnection(DbPathName)
    Set this.DbConn = DbConn
    Set DbConnInit = DbConn
End Function


Public Property Get DbStmtGet(Optional ByVal StmtName As String = "main") As SQLiteCStatement
    Set DbStmtGet = this.DbConn.DbStmt(StmtName)
End Property


Public Function CreateStatement(ByVal StmtName As String) As SQLiteCStatement
    Set CreateStatement = this.DbConn.CreateStatement(StmtName)
End Function


'''' Reference: https://www.sqlite.org/c3ref/backup_finish.html
'''' Reference: https://www.sqlite.org/backup.html
'''' Reference: https://www.sqlite.org/lang_vacuum.html#vacuuminto
''''
'@Description "Performs online database backup."
Public Function DupOnlineFull(ByVal DbConnDst As SQLiteCConnection, _
                     Optional ByVal DstAlias As String = "main", _
                     Optional ByVal DbConnSrc As SQLiteCConnection, _
                     Optional ByVal SrcAlias As String = "main") As SQLiteResultCodes
Attribute DupOnlineFull.VB_Description = "Performs online database backup."
    Guard.NullReference DbConnSrc
    If DbConnDst Is Nothing Then Set DbConnDst = DbConnSrc
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
    
RESULT_CODE:
    ResultCode = DbConnDst.ErrInfoRetrieve
    DupOnlineFull = ResultCode
    Exit Function
End Function
