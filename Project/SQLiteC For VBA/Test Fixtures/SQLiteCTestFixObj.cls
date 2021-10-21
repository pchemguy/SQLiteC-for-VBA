VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCTestFixObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides objects for tests"
'@Folder "SQLiteC For VBA.Test Fixtures"
'@ModuleDescription "Provides objects for tests"
'@PredeclaredId
'@IgnoreModule MoveFieldCloserToUsage, ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Private Const LITE_LIB As String = "SQLiteCforVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TSQLiteCConnectionTestFixtures
    DllMan As DllManager
End Type
Private this As TSQLiteCConnectionTestFixtures


Public Function GetDefaultDBM() As SQLiteC
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim dbm As SQLiteC
    '''' Using default library names hardcoded in the SQLiteC constructor.
    Set dbm = SQLiteC(DllPath)
    If dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteC instance."
    Set GetDefaultDBM = dbm
End Function

Public Function GetConnection(ByVal DbPathName As String) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = GetDefaultDBM()
    '''' Keep a reference to the DllManager to prevent library unloading due to object destruction
    Set this.DllMan = dbm.DllMan
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTestFixtures", "Failed to create an SQLiteCConnection instance."
    Set GetConnection = DbConn
End Function

Public Function GetConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set GetConnDbRegular = GetConnection(DbPathName)
End Function

Public Function GetConnDbReadOnlyAttr() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & "ReadOnlyAttr.db"
    Set GetConnDbReadOnlyAttr = GetConnection(DbPathName)
End Function

Public Function GetConnDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetConnDbMemory = GetConnection(DbPathName)
End Function

Public Function GetConnDbAnon() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set GetConnDbAnon = GetConnection(DbPathName)
End Function

Public Function GetConnDbTemp() As SQLiteCConnection
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    Dim DbFilePath As String
    DbFilePath = Environ("TEMP")
    Dim DbFileName As String
    DbFileName = Replace(Replace(Replace(Now(), "/", "-"), " ", "_"), ":", "-")
    Dim DbFileExt As String
    DbFileExt = ".db"
    Dim DbPathName As String
    DbPathName = fso.BuildPath(DbFilePath, DbFileName) & DbFileExt
    Set GetConnDbTemp = GetConnection(DbPathName)
End Function

Public Function GetConnDbInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set GetConnDbInvalidPath = GetConnection(DbPathName)
End Function

Public Function GetConnDbMemoryWithTable() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set GetConnDbMemoryWithTable = GetConnection(DbPathName)
    GetConnDbMemoryWithTable.ExecuteNonQueryPlain SQLiteCTestFixSQL.CREATETableITRB
End Function

Public Function CreateFunctionsTableWithData(ByVal dbc As SQLiteCConnection) As Long
    Dim SQLQuery As String
    SQLQuery = SQLiteCTestFixSQL.CreateFunctionsTableWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    CreateFunctionsTableWithData = AffectedRows
End Function

'@Description "Generates byte array from either multiple single byte arguments or a single string of ASCII letters"
Public Function ByteArray(ParamArray Values() As Variant) As Byte()
Attribute ByteArray.VB_Description = "Generates byte array from either multiple single byte arguments or a single string of ASCII letters"
    Dim ValuesLocal() As Variant
    ValuesLocal = UnfoldParamArray(Values)
    Dim Buffer() As Byte
    Dim Size As Long
    Dim ByteIndex As Long
    If UBound(ValuesLocal) > 0 Or VarType(ValuesLocal(0)) = vbByte Then
        Size = UBound(ValuesLocal)
        ReDim Buffer(0 To Size)
        Dim Element As Variant
        For ByteIndex = 0 To Size
            Element = ValuesLocal(ByteIndex)
            Select Case VarType(Element)
                Case vbByte
                    Buffer(ByteIndex) = Element
                Case vbInteger, vbLong
                    If Element > 255 Then GoTo INVALID_INPUT:
                    Buffer(ByteIndex) = CByte(Element)
                Case vbString
                    If Len(Element) <> 1 Then GoTo INVALID_INPUT:
                    Buffer(ByteIndex) = Asc(Element)
                Case Else
                    GoTo INVALID_INPUT:
            End Select
        Next ByteIndex
    ElseIf UBound(ValuesLocal) = 0 And VarType(ValuesLocal(0)) = vbString Then
        Size = Len(ValuesLocal(0)) - 1
        ReDim Buffer(0 To Size)
        For ByteIndex = 0 To Size
            Buffer(ByteIndex) = Asc(Mid(ValuesLocal(0), ByteIndex + 1, 1))
        Next ByteIndex
    Else
        GoTo INVALID_INPUT:
    End If
    
    ByteArray = Buffer
    Exit Function
    
INVALID_INPUT:
    Err.Raise ErrNo.TypeMismatchErr, "SQLiteCTestFixObj", _
              "Invalid input for byte array. Element #" & _
              CStr(ByteIndex + 1) & " <<" & CStr(Element) & ">>"
End Function

'@Description "Generates VBA string from ASCII byte array"
Public Function AsciiByteArrayToString(ByRef AsciiBytes() As Byte) As String
Attribute AsciiByteArrayToString.VB_Description = "Generates VBA string from ASCII byte array"
    AsciiByteArrayToString = StrConv(AsciiBytes, vbUnicode)
End Function

'@Description "Generates Dictionary from separate arrays of keys and values"
Public Function KeysValuesToDict(ByVal Keys As Variant, ByVal Values As Variant) As Scripting.Dictionary
Attribute KeysValuesToDict.VB_Description = "Generates Dictionary from separate arrays of keys and values"
    If Not (IsArray(Keys) And IsArray(Values)) Then
        Err.Raise ErrNo.InvalidParameterErr, "SQLiteCTestFixObj", "Expected two arrays"
    End If
    If LBound(Keys) <> LBound(Values) Or UBound(Keys) <> UBound(Values) Then
        Err.Raise ErrNo.IncompatibleArraysErr, "SQLiteCTestFixObj", "Incompatible arrays"
    End If
    Dim Items As Scripting.Dictionary
    Set Items = New Scripting.Dictionary
    Items.CompareMode = TextCompare
    Dim ItemIndex As Long
    For ItemIndex = LBound(Keys) To UBound(Keys)
        Items.Item(Keys(ItemIndex)) = Values(ItemIndex)
    Next ItemIndex
    Set KeysValuesToDict = Items
End Function

Public Function ByteToHex(ByVal ByteCode As Byte) As String
    Dim HexDigitMap As Variant
    HexDigitMap = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
    ByteToHex = HexDigitMap(ByteCode \ 16) & HexDigitMap(ByteCode Mod 16)
End Function

Public Function XorElements(ByVal Elements As Variant) As Long
    If Not IsArray(Elements) Then Err.Raise ErrNo.InvalidParameterErr, "Expected an array"
    Dim ElementIndex As Long
    Dim Result As Long
    Result = 0
    For ElementIndex = LBound(Elements) To UBound(Elements)
        Result = Result Xor Elements(ElementIndex)
    Next ElementIndex
    XorElements = Result
End Function
