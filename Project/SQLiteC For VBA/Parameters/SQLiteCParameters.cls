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
'''' https://www.sqlite.org/c3ref/bind_parameter_index.html
'''' https://www.sqlite.org/c3ref/bind_parameter_count.html
'''' https://www.sqlite.org/c3ref/bind_parameter_name.html
''''======================================================================''''

Option Explicit

Public Enum SQLiteMemoryManagement
    SQLITE_STATIC = 0&
    SQLITE_TRANSIENT = -1&
End Enum

#If VBA7 Then
' 1-based indices
Private Declare PtrSafe Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramName As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Double) As Long
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As LongLong) As Long
Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
#Else
' 1-based indices
Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamName As Long) As Long
Private Declare Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal nBytes As Long) As Long
Private Declare Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Double) As Long
Private Declare Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Currency) As Long ' UNTESTED ....?
Private Declare Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pSqlite3Value As Long) As Long
Private Declare Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As Long) As Long
#End If

Private Type TSQLiteCParameters
    DbStmt As SQLiteCStatement
End Type
Private this As TSQLiteCParameters


'@DefaultMember
Public Function Create(ByVal DbStmt As SQLiteCStatement) As SQLiteCParameters
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCParameters
    Set Instance = New SQLiteCParameters
    Instance.Init DbStmt
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbStmt As SQLiteCStatement)
    Guard.NullReference DbStmt
    Set this.DbStmt = DbStmt
End Sub


Public Property Get Count() As Long
    Count = sqlite3_bind_parameter_count(this.DbStmt.StmtHandle)
End Property


Public Function IndexFromName(ByVal ParamName As String) As Long
    Guard.EmptyString ParamName
    '@Ignore UseMeaningfulName
    Dim ParamNameUTF8() As Byte
    ParamNameUTF8 = UTFlib.UTF8BytesFromStr(ParamName)
    #If VBA7 Then
        Dim ParamNamePtr As LongPtr
    #Else
        Dim ParamNamePtr As Long
    #End If
    ParamNamePtr = VarPtr(ParamNameUTF8(0))
    IndexFromName = sqlite3_bind_parameter_index(this.DbStmt.StmtHandle, ParamNamePtr)
End Function


Public Function NameFromIndex(ByVal ParamIndex As Long) As String
    Guard.ExpressionErr ParamIndex > 0, InvalidParameterErr, "SQLiteCParameters", _
        "ParamIndex must be a positive integer"
    NameFromIndex = UTFlib.StrFromUTF8Ptr(sqlite3_bind_parameter_name( _
        this.DbStmt.StmtHandle, ParamIndex))
End Function


Public Function BindOne(ByVal ParamIndex As Long, ByVal Value As Variant) As Boolean
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    Dim Length As Long
    Dim ResultCode As SQLiteResultCodes
    ResultCode = SQLITE_ERROR
    Select Case VarType(Value)
        Case vbInteger, vbLong
            ResultCode = sqlite3_bind_int(hStmt, ParamIndex, Value)
        Case vbSingle, vbDouble
            ResultCode = sqlite3_bind_double(hStmt, ParamIndex, Value)
        Case vbString
            Length = Len(Value) * 2
            ResultCode = sqlite3_bind_text16(hStmt, ParamIndex, StrPtr(Value), Length, SQLITE_TRANSIENT)
        Case vbEmpty, vbNull
            ResultCode = sqlite3_bind_null(hStmt, ParamIndex)
        Case vbArray + vbByte
            Length = UBound(Value) - LBound(Value) + 1
            ResultCode = sqlite3_bind_blob(hStmt, ParamIndex, VarPtr(Value(0)), Length, SQLITE_TRANSIENT)
        Case vbObject, vbError, vbDataObject, vbUserDefinedType
            Debug.Print "Unsupported value type: " & CStr(VarType(Value))
            ResultCode = SQLITE_ERROR
        Case Is < vbArray
            Debug.Print "Unsupported value converted to String. Type: " & CStr(VarType(Value))
            Dim StrValue As String
            StrValue = CStr(Value)
            Length = Len(StrValue) * 2
            ResultCode = sqlite3_bind_text16(hStmt, ParamIndex, StrPtr(StrValue), Length, SQLITE_TRANSIENT)
        Case Is >= vbArray
            Debug.Print "Unsupported array type: " & CStr(VarType(Value))
            ResultCode = SQLITE_ERROR
    End Select
    If ResultCode > SQLITE_ERROR Then
        With this.DbStmt.DbConn.ErrorInfo
            .GetErr
            .PrintErr
        End With
    End If
    BindOne = (ResultCode = SQLITE_OK)
End Function


Public Function BindSeq(ParamArray ParamValues() As Variant) As Boolean
    Dim Values() As Variant
    Values = UnfoldParamArray(ParamValues)
    Dim ValueCount As Long
    ValueCount = UBound(Values, 1) - LBound(Values, 1) + 1
    Dim ParamCount As Long
    ParamCount = Count
    If ValueCount <> ParamCount Then
        Debug.Print "Wrong number of parameter values:" & vbNewLine & _
            "Expected: " & CStr(ParamCount) & vbNewLine & "Received: " & _
            CStr(ValueCount)
        BindSeq = False
        Exit Function
    End If
    Dim Result As Boolean
    Result = True
    Dim ParamIndex As Long
    For ParamIndex = 1 To ParamCount
        Result = Result And BindOne(ParamIndex, Values(ParamIndex - 1))
    Next ParamIndex
    Dim ResultCode As SQLiteResultCodes
    If Not Result Then
        ResultCode = sqlite3_clear_bindings(this.DbStmt.StmtHandle)
        If ResultCode <> SQLITE_OK Then
            With this.DbStmt.DbConn.ErrorInfo
                .GetErr
                .PrintErr
            End With
        End If
    End If
    BindSeq = Result
End Function


Public Function BindMap(ByVal ParamValues As Scripting.Dictionary) As Boolean
    Guard.NullReference ParamValues
    Dim Result As Boolean
    Result = True
    Dim ParamIndex As Long
    Dim ParamName As Variant
    For Each ParamName In ParamValues.Keys
        ParamIndex = IndexFromName(ParamName)
        Result = Result And BindOne(ParamIndex, ParamValues(ParamName))
    Next ParamName
    Dim ResultCode As SQLiteResultCodes
    If Not Result Then
        ResultCode = sqlite3_clear_bindings(this.DbStmt.StmtHandle)
        If ResultCode <> SQLITE_OK Then
            With this.DbStmt.DbConn.ErrorInfo
                .GetErr
                .PrintErr
            End With
        End If
    End If
    BindMap = Result
End Function
