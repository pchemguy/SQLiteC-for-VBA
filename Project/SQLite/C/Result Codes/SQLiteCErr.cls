VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Handles result codes"
'@Folder "SQLite.C.Result Codes"
'@ModuleDescription "Handles result codes"
'@PredeclaredId
'@Exposed

''''======================================================================''''
'''' Error Codes
'''' https://www.sqlite.org/c3ref/errcode.html
''''======================================================================''''

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteResultCodes) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
#Else
Private Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteResultCodes) As Long ' PtrUtf8String
Private Declare Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
#End If

Private Type TSQLiteCErr
    DbConn As SQLiteCConnection
    ErrorCode As SQLiteResultCodes
    ErrorCodeEx As SQLiteResultCodes
    ErrorMessage As String
    ErrorString As String
End Type
Private this As TSQLiteCErr


'@DefaultMember
Public Function Create(ByVal DbConn As SQLiteCConnection) As SQLiteCErr
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCErr
    Set Instance = New SQLiteCErr
    Instance.Init DbConn
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbConn As SQLiteCConnection)
    Guard.NullReference DbConn
    Set this.DbConn = DbConn
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

Friend Sub Cleanup()
    Set this.DbConn = Nothing
End Sub

Public Property Get ErrorCode() As SQLiteResultCodes
    ErrorCode = this.ErrorCode
End Property

Public Property Get ErrorCodeEx() As SQLiteResultCodes
    ErrorCodeEx = this.ErrorCodeEx
End Property

Public Property Get ErrorCodeName() As String
    ErrorCodeName = SQLiteCRC.CodeToName(this.ErrorCode)
End Property

Public Property Get ErrorCodeExName() As String
    ErrorCodeExName = SQLiteCRC.CodeToName(this.ErrorCodeEx)
End Property

Public Property Get ErrorName() As String
    ErrorName = ErrorCodeExName
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = this.ErrorMessage
End Property

Public Property Get ErrorString() As String
    ErrorString = this.ErrorString
End Property

Public Sub GetErr()
    #If VBA7 Then
        Dim DbHandle As LongPtr
    #Else
        Dim DbHandle As Long
    #End If
    DbHandle = this.DbConn.DbHandle
    If DbHandle = 0 Then Err.Raise ErrNo.ConnectionNotOpenedErr, _
        "SQLiteCErr", "Database connection is not opened!"
    With this
        .ErrorCode = sqlite3_errcode(DbHandle)
        .ErrorCodeEx = sqlite3_extended_errcode(DbHandle)
        .ErrorMessage = UTFlib.StrFromUTF8Ptr(sqlite3_errmsg(DbHandle))
        .ErrorString = UTFlib.StrFromUTF8Ptr(sqlite3_errstr(.ErrorCode))
    End With
End Sub

Public Sub PrintErr()
    With this
        Logger.Logg "========== Error Details ========== " & vbNewLine & _
                    "ErrCode:          " & CStr(.ErrorCode) & vbNewLine & _
                    "ErrCode Extended: " & CStr(.ErrorCodeEx) & vbNewLine & _
                    "Error Name:       " & ErrorName & vbNewLine & _
                    "ErrStr:           " & .ErrorString & vbNewLine & _
                    "ErrMsg:           " & .ErrorMessage & vbNewLine _
                    , , DEBUGLEVEL_ERROR
    End With
End Sub


'#If VBA7 Then
'Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf16String
'#Else
'Private Declare Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf16String
'#End If
