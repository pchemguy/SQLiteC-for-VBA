VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Wraps SQLite3 statement commands."
'@Folder "SQLiteC For VBA.Statement"
'@ModuleDescription "Wraps SQLite3 statement commands."
'@PredeclaredId
'@Exposed

''''======================================================================''''
'''' Statement
'''' https://www.sqlite.org/c3ref/prepare.html
'''' https://www.sqlite.org/c3ref/reset.html
'''' https://www.sqlite.org/c3ref/finalize.html
''''======================================================================''''

Option Explicit

Const OutOfMemoryErr As Long = 7&

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_prepare16_v2 Lib "SQLite3" (ByVal hDb As LongPtr, ByVal pwsSql As LongPtr, _
    ByVal nSqlLength As Long, ByRef hStmt As LongPtr, ByVal ppwsTailOut As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_step Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As LongPtr) As SQLiteResultCodes
Private Declare PtrSafe Function sqlite3_sql Lib "SQLite3" (ByVal hStmt As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_expanded_sql Lib "SQLite3" (ByVal hStmt As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_normalized_sql Lib "SQLite3" (ByVal hStmt As Long) As LongPtr ' PtrString
Private Declare PtrSafe Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As LongPtr)
#Else
Private Declare Function sqlite3_prepare16_v2 Lib "SQLite3" (ByVal hDb As Long, ByVal pwsSql As Long, _
    ByVal nSqlLength As Long, ByRef hStmt As Long, ByVal ppwsTailOut As Long) As SQLiteResultCodes
Private Declare Function sqlite3_step Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_reset Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_finalize Lib "SQLite3" (ByVal hStmt As Long) As SQLiteResultCodes
Private Declare Function sqlite3_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Function sqlite3_expanded_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Function sqlite3_normalized_sql Lib "SQLite3" (ByVal hStmt As Long) As Long ' PtrString
Private Declare Sub sqlite3_free Lib "SQLite3" (ByVal MemoryPtr As Long)
#End If

Private Type TSQLiteCStatement
    DbConn As SQLiteCConnection
    DbCmd As SQLiteCCommand
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


Private Sub Class_Terminate()
    If this.StmtHandle <> 0 Then Finalize
End Sub


Public Property Get DbConn() As SQLiteCConnection
    Set DbConn = this.DbConn
End Property


#If VBA7 Then
Public Property Get StmtHandle() As LongPtr
#Else
Public Property Get StmtHandle() As Long
#End If
    StmtHandle = this.StmtHandle
End Property


Public Sub ErrInfoRetrieve()
    With this.DbConn.ErrorInfo
        .GetErr
        .PrintErr
    End With
End Sub


Public Property Get SQLQueryExpanded() As String
    #If VBA7 Then
        Dim SQLQueryPtr As LongPtr
    #Else
        Dim SQLQueryPtr As Long
    #End If
    SQLQueryPtr = sqlite3_expanded_sql(this.StmtHandle)
    If SQLQueryPtr = 0 Then
        Err.Raise OutOfMemoryErr, "SQLiteCStatement/SQLQueryExpanded", _
                  "SQLite DLL returned NULL pointer"
    End If
    Dim SQLQueryText As String
    SQLQueryText = UTFlib.StrFromUTF8Ptr(SQLQueryPtr)
    sqlite3_free SQLQueryPtr
    SQLQueryExpanded = SQLQueryText
End Property


Public Function SQLQuery(Optional ByVal Normalized As Boolean = False) As String
    #If VBA7 Then
        Dim SQLQueryPtr As LongPtr
    #Else
        Dim SQLQueryPtr As Long
    #End If
    If Normalized Then
        SQLQueryPtr = sqlite3_normalized_sql(this.StmtHandle)
    Else
        SQLQueryPtr = sqlite3_sql(this.StmtHandle)
    End If
    If SQLQueryPtr = 0 Then
        Err.Raise OutOfMemoryErr, "SQLiteCStatement/SQLQuery", _
                  "SQLite DLL returned NULL pointer"
    End If
    SQLQuery = UTFlib.StrFromUTF8Ptr(SQLQueryPtr)
End Function


'@Ignore UseMeaningfulName
Public Function Prepare16V2(ByVal SQLQuery As String) As Boolean
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_prepare16_v2(this.DbConn.DbHandle, StrPtr(SQLQuery), _
                                      Len(SQLQuery) * 2 + 2, this.StmtHandle, 0)
    If ResultCode = SQLITE_OK Then
        Prepare16V2 = True
    Else
        ErrInfoRetrieve
        Prepare16V2 = False
    End If
End Function


Public Function Reset() As Boolean
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_reset(this.StmtHandle)
    If ResultCode = SQLITE_OK Then
        Reset = True
    Else
        ErrInfoRetrieve
        Reset = False
    End If
End Function


Public Function Finalize() As Boolean
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_finalize(this.StmtHandle)
    If ResultCode = SQLITE_OK Then
        this.StmtHandle = 0
        Finalize = True
    Else
        ErrInfoRetrieve
        Finalize = False
    End If
End Function
