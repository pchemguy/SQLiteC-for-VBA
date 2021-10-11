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

''''======================================================================''''
'''' Overview
'''' https://www.sqlite.org/cintro.html
''''======================================================================''''

Option Explicit

Private Type TSQLiteC
    DllMan As DllManager
    DbConn As SQLiteCConnection
    DbStmt As SQLiteCStatement
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


Public Property Get DbStmtGet() As SQLiteCStatement
    Set DbStmtGet = this.DbStmt
End Property


Public Function DbStmtInit() As SQLiteCStatement
    Guard.NullReference this.DbConn
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = SQLiteCStatement(this.DbConn)
    Set this.DbStmt = DbStmt
    Set DbStmtInit = DbStmt
End Function
