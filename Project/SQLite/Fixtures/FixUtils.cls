VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides convenience utilities"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides convenience utilities"
'@PredeclaredId
Option Explicit


'@Description "Generates byte array from either multiple single byte arguments or a single string of ASCII letters"
Public Function ByteArray(ParamArray Values() As Variant) As Byte()
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
    AsciiByteArrayToString = StrConv(AsciiBytes, vbUnicode)
End Function


'@Description "Generates Dictionary from separate arrays of keys and values"
Public Function KeysValuesToDict(ByVal Keys As Variant, ByVal Values As Variant) As Scripting.Dictionary
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
