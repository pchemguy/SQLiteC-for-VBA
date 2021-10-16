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
'@IgnoreModule IndexedDefaultMemberAccess

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
Private Declare PtrSafe Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Double) As Long
Private Declare PtrSafe Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As Long) As Long
Private Declare PtrSafe Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal Value As LongLong) As Long
Private Declare PtrSafe Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pswValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_clear_bindings Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
#Else
' 1-based indices
Private Declare Function sqlite3_bind_parameter_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_bind_parameter_name Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long) As Long
Private Declare Function sqlite3_bind_parameter_index Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamName As Long) As Long
Private Declare Function sqlite3_bind_null Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long) As Long
Private Declare Function sqlite3_bind_blob Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
Private Declare Function sqlite3_bind_double Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Double) As Long
Private Declare Function sqlite3_bind_int Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Long) As Long
Private Declare Function sqlite3_bind_int64 Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal Value As Currency) As Long '
Private Declare Function sqlite3_bind_text16 Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pswValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
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


Private Sub Class_Terminate()
    BindClear
End Sub


'''' Reference: https://www.sqlite.org/c3ref/bind_parameter_count.html
''''
'@Description "Retrieves parameter count."
Public Property Get Count() As Long
Attribute Count.VB_Description = "Retrieves parameter count."
    Count = sqlite3_bind_parameter_count(this.DbStmt.StmtHandle)
End Property


'''' Reference: https://www.sqlite.org/c3ref/bind_parameter_index.html
''''
'@Description "Retrieves named parameter index."
Public Function IndexFromName(ByVal ParamName As String) As Long
Attribute IndexFromName.VB_Description = "Retrieves named parameter index."
    Guard.EmptyString ParamName
    '@Ignore UseMeaningfulName
    Dim ParamNameUTF8() As Byte
    ParamNameUTF8 = UTFlib.UTF8BytesFromStr(ParamName)
    IndexFromName = sqlite3_bind_parameter_index(this.DbStmt.StmtHandle, _
                                                 VarPtr(ParamNameUTF8(0)))
End Function


'''' Reference: https://www.sqlite.org/c3ref/bind_parameter_name.html
''''
'@Ignore ProcedureNotUsed
'@Description "Retrieves named parameter name."
Public Function NameFromIndex(ByVal ParamIndex As Long) As String
Attribute NameFromIndex.VB_Description = "Retrieves named parameter name."
    Guard.ExpressionErr ParamIndex > 0, InvalidParameterErr, "SQLiteCParameters", _
        "ParamIndex must be a positive integer"
    NameFromIndex = UTFlib.StrFromUTF8Ptr(sqlite3_bind_parameter_name( _
        this.DbStmt.StmtHandle, ParamIndex))
End Function


'''' Reference: https://www.sqlite.org/c3ref/bind_blob.html
''''
'@Description "Binds a single parameter. A particular interface is selected based on the value type."
Public Function BindOne(ByVal ParamIndex As Long, ByVal Value As Variant) As SQLiteResultCodes
Attribute BindOne.VB_Description = "Binds a single parameter. A particular interface is selected based on the value type."
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = this.DbStmt.StmtHandle
    Dim Length As Long
    Dim ResultCode As SQLiteResultCodes
    Select Case VarType(Value)
        Case vbByte, vbInteger, vbLong
            ResultCode = sqlite3_bind_int(hStmt, ParamIndex, Value)
        Case vbCurrency, vbLongLong
            ResultCode = sqlite3_bind_int64(hStmt, ParamIndex, Value)
        Case vbSingle, vbDouble
            ResultCode = sqlite3_bind_double(hStmt, ParamIndex, Value)
        Case vbString
            Length = Len(Value) * 2
            ResultCode = sqlite3_bind_text16(hStmt, ParamIndex, StrPtr(Value), Length, SQLITE_TRANSIENT)
        Case vbBoolean
            ResultCode = sqlite3_bind_int(hStmt, ParamIndex, Abs(Value))
        Case vbEmpty, vbNull
            ResultCode = sqlite3_bind_null(hStmt, ParamIndex)
        Case vbArray + vbByte
            Length = UBound(Value) - LBound(Value) + 1
            ResultCode = sqlite3_bind_blob(hStmt, ParamIndex, VarPtr(Value(0)), Length, SQLITE_TRANSIENT)
        Case vbDecimal
            Dim DecValue As String
            DecValue = CStr(Value)
            Length = Len(DecValue) * 2
            ResultCode = sqlite3_bind_text16(hStmt, ParamIndex, StrPtr(DecValue), Length, SQLITE_TRANSIENT)
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
    If ResultCode > SQLITE_ERROR Then ResultCode = this.DbStmt.DbConn.ErrInfoRetrieve
    BindOne = ResultCode
End Function


'@Description "Binds a set of parameter values provided as an array. Must match the number of available parameters."
Public Function BindArray(ByVal ValueArr As Variant) As SQLiteResultCodes
Attribute BindArray.VB_Description = "Binds a set of parameter values provided as an array. Must match the number of available parameters."
    Guard.ExpressionErr VarType(ValueArr) >= vbArray, ErrNo.InvalidParameterErr, _
            "SQLiteCParameters", "Expected an array parameter, received VarType: " _
            & CStr(VarType(ValueArr))
    Dim ValueCount As Long
    ValueCount = UBound(ValueArr, 1) - LBound(ValueArr, 1) + 1
    Dim ParamCount As Long
    ParamCount = Count
    If ValueCount <> ParamCount Then
        Debug.Print "Wrong number of parameter values:" & vbNewLine & _
            "Expected: " & CStr(ParamCount) & vbNewLine & "Received: " & _
            CStr(ValueCount)
        BindArray = SQLITE_ERROR
        Exit Function
    End If
    Dim ResultCode As SQLiteResultCodes
    Dim ParamIndex As Long
    For ParamIndex = 1 To ParamCount
        ResultCode = BindOne(ParamIndex, ValueArr(ParamIndex - 1))
        If ResultCode <> SQLITE_OK Then
            BindArray = ResultCode
            Exit Function
        End If
    Next ParamIndex
    BindArray = ResultCode
End Function


'@Description "Binds a set of parameter values provided as Name->Value map. Only parameters with matching names are bound."
Public Function BindDict(ByVal ValueDict As Scripting.Dictionary) As SQLiteResultCodes
Attribute BindDict.VB_Description = "Binds a set of parameter values provided as Name->Value map. Only parameters with matching names are bound."
    Guard.NullReference ValueDict
    Dim ResultCode As SQLiteResultCodes
    Dim ParamIndex As Long
    Dim ParamName As Variant
    For Each ParamName In ValueDict.Keys
        ParamIndex = IndexFromName(ParamName)
        If ParamIndex > 0 Then
            ResultCode = BindOne(ParamIndex, ValueDict(ParamName))
            If ResultCode <> SQLITE_OK Then
                BindDict = ResultCode
                BindClear
                Exit Function
            End If
        End If
    Next ParamName
    BindDict = ResultCode
End Function


Public Function BindDictOrArray(ByVal ParamValues As Variant) As SQLiteResultCodes
    If this.DbStmt.StmtHandle = 0 Then Err.Raise StatementNotPreparedErr, _
        "SQLiteCParameters", "Statement is not prepared."
    If TypeName(ParamValues) = "Dictionary" Then
        BindDictOrArray = BindDict(ParamValues)
    Else
        BindDictOrArray = BindArray(ParamValues)
    End If
End Function


Public Sub BindClear()
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_clear_bindings(this.DbStmt.StmtHandle)
    '@Ignore FunctionReturnValueDiscarded
    If ResultCode <> SQLITE_OK Then this.DbStmt.DbConn.ErrInfoRetrieve
End Sub


'#If VBA7 Then
'' 1-based indices
'Private Declare PtrSafe Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal nBytes As Long) As Long
'Private Declare PtrSafe Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal psValue As LongPtr, ByVal nBytes As Long, ByVal pfDelete As LongPtr) As Long
'Private Declare PtrSafe Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal paramIndex As Long, ByVal pSqlite3Value As LongPtr) As Long
'#Else
'' 1-based indices
'Private Declare Function sqlite3_bind_zeroblob Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal nBytes As Long) As Long
'Private Declare Function sqlite3_bind_text Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal psValue As Long, ByVal nBytes As Long, ByVal pfDelete As Long) As Long
'Private Declare Function sqlite3_bind_value Lib "SQLite3" (ByVal hStmt As Long, ByVal ParamIndex As Long, ByVal pSqlite3Value As Long) As Long
'#End If
