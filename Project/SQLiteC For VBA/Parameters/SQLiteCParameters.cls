VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Handles query parameters"
'@Folder "SQLiteC For VBA.Parameters"
'@ModuleDescription "Handles query parameters"
'@PredeclaredId
'@Exposed

''''======================================================================''''
'''' Parameters
'''' https://www.sqlite.org/c3ref/bind_blob.html
''''
''''======================================================================''''

Option Explicit

#If VBA7 Then
' Statements
Private Declare PtrSafe Function sqlite3_exec Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal zSql As LongPtr, ByVal Callback As LongPtr, _
     ByVal pArg As LongPtr, ByRef pzErrMsg As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, ByVal nSqlLength As Long, _
     ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As LongPtr)
#Else
' Statements
Private Declare Function sqlite3_exec Lib "SQLite3" _
    (ByVal hDb As Long, ByVal zSql As Long, ByVal Callback As Long, _
     ByVal pArg As Long, ByRef pzErrMsg As Long) As SQLiteResultCodes
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3" _
    (ByVal hDb As Long, ByVal pwsSql As Long, ByVal nSqlLength As Long, _
     ByRef hStmt As Long, ByVal ppwsTailOut As Long) As SQLiteResultCodes
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As Long)
#End If

Private Type TSQLiteCStatement
    DbConn As SQLiteCConnection
    #If VBA7 Then
        StmtHandle As LongPtr
    #Else
        StmtHandle As Long
    #End If
End Type
Private this As TSQLiteCStatement


'@DefaultMember
Public Function Create(ByVal DbConn As SQLiteCConnection) As SQLiteCStatement
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCStatement
    Set Instance = New SQLiteCStatement
    Instance.Init DbConn
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbConn As SQLiteCConnection)
    Guard.NullReference DbConn
    Set this.DbConn = DbConn
End Sub
