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
'''' https://www.sqlite.org/c3ref/step.html
'''' https://www.sqlite.org/c3ref/reset.html
'''' https://www.sqlite.org/c3ref/finalize.html
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


Private Sub Class_Terminate()
    If this.StmtHandle <> 0 Then Finalize
End Sub


#If VBA7 Then
Public Property Get StmtHandle() As LongPtr
#Else
Public Property Get StmtHandle() As Long
#End If
    StmtHandle = this.StmtHandle
End Property


'@Ignore UseMeaningfulName
Public Function Prepare16V2(ByVal SQLQuery As String) As Boolean
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_prepare16_v2(this.DbConn.DbHandle, StrPtr(SQLQuery), _
                                      Len(SQLQuery) * 2 + 2, this.StmtHandle, 0)
    If ResultCode = SQLITE_OK Then
        Prepare16V2 = True
    Else
        With this.DbConn.ErrorInfo
            .GetErr
            .PrintErr
        End With
        Prepare16V2 = False
    End If
End Function


Public Function Reset() As Boolean
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_reset(this.StmtHandle)
    If ResultCode = SQLITE_OK Then
        Reset = True
    Else
        With this.DbConn.ErrorInfo
            .GetErr
            .PrintErr
        End With
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
        With this.DbConn.ErrorInfo
            .GetErr
            .PrintErr
        End With
        Finalize = False
    End If
End Function


Public Function ExecuteNonQuery(ByVal SQLQuery As String) As Long
    Guard.EmptyString SQLQuery
    
    Dim ChangesCount As Long
    ChangesCount = this.DbConn.ChangesCount(True)
    
    Dim SQLQueryUTF8() As Byte
    SQLQueryUTF8 = UTFlib.UTF8BytesFromStr(SQLQuery)
    #If VBA7 Then
        Dim SQLQueryUTF8Ptr As LongPtr
        Dim CallbackRowsPtr As LongPtr
        Dim ArgPtr As LongPtr
        Dim ErrMsgPtr As LongPtr
    #Else
        Dim SQLQueryUTF8Ptr As Long
        Dim CallbackRowsPtr As Long
        Dim ArgPtr As Long
        Dim ErrMsgUTF8Ptr As Long
    #End If
    SQLQueryUTF8Ptr = VarPtr(SQLQueryUTF8(0))
    CallbackRowsPtr = 0
    ArgPtr = 0
    ErrMsgUTF8Ptr = 0
    
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_exec(this.DbConn.DbHandle, SQLQueryUTF8Ptr, _
                              CallbackRowsPtr, ArgPtr, ErrMsgUTF8Ptr)
    
    If ResultCode <> SQLITE_OK Then
        With this.DbConn.ErrorInfo
            .GetErr
            .PrintErr
        End With
        Dim ErrMsg As String
        ErrMsg = UTFlib.StrFromUTF8Ptr(ErrMsgUTF8Ptr)
        Debug.Print ErrMsg
        ExecuteNonQuery = -1&
    Else
        ChangesCount = this.DbConn.ChangesCount(True) - ChangesCount
        ExecuteNonQuery = ChangesCount
    End If
    sqlite3_free ErrMsgUTF8Ptr
End Function
