Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
'@IgnoreModule ProcedureNotUsed, FunctionReturnValueDiscarded
Option Explicit

'@IgnoreModule MoveFieldCloserToUsage
Private lastID As Double


Public Function GetTimeStampMs() As String
    '''' On Windows, the Timer resolution is subsecond, the fractional part (the four characters at the end
    '''' given the format) is concatenated with DateTime. It appears that the Windows' high precision time
    '''' source available via API yields garbage for the fractional part.
    GetTimeStampMs = Format$(Now, "yyyy-MM-dd HH:mm:ss") & Right$(Format$(Timer, "#0.000"), 4)
End Function


'''' The number of seconds since the Epoch is multiplied by 10^4 to bring the first
'''' four fractional places in Timer value into the whole part before trancation.
'''' Long on a 32bit machine does not provide sufficient number of digits,
'''' so returning double. Alternatively, a Currency type could be used.
Public Function GenerateSerialID() As Double
    Dim newID As Double
    Dim secTillLastMidnight As Double
    secTillLastMidnight = CDbl(DateDiff("s", DateSerial(1970, 1, 1), Date))
    newID = Fix((secTillLastMidnight + Timer) * 10 ^ 4)
    If newID > lastID Then
        lastID = newID
    Else
        lastID = lastID + 1
    End If
    GenerateSerialID = lastID
    'GetSerialID = Fix((CDbl(Date) * 100000# + CDbl(Timer) / 8.64))
End Function


'''' Unfolds a ParamArray argument
''''
'''' When sub/function captures a list of arguments in a ParamArray and passes it
'''' to the next routine expecting a list of arguments, the second routine receives
'''' a 2D array instead of 1D with the outer dimension having a single element.
'''' This function check the arguments and unfolds the outer dimesion as necessary.
'''' Any function accepting a ParamArray argument should be able to use it.
''''
'''' Args:
''''   ParamArrayArg:
''''     An argument that was relayed as ParamArray twice sequentially.
''''     If all of the following conditions are satisfied:
''''       - ParamArrayArg is a 1D array
''''       - UBound(ParamArrayArg, 1) = LBound(ParamArrayArg, 1) = 0
''''       - ParamArrayArg(0) is a 1D 0-based array
''''     the inner array is extracted and returned.
''''
'''' Returns:
''''   ParamArrayArg(0), if unfolding is necessary
''''   ParamArrayArg, if ParamArrayArg is an array, but not all conditions are satisfied
''''
'''' Raises:
''''   ErrNo.ExpectedArrayErr:
''''     If ParamArrayArg is not an array.
''''
'''' Examples:
''''   Raises error:
''''     >>> ?UnfoldParamArray("A")
''''     Raises "Expected array" error
''''
''''   Returns as is without unfolding:
''''     >>> UnfoldParamArray(Array("A", "B", "C"))
''''     Array("A", "B", "C")
''''     >>> ?Join(UnfoldParamArray(Array("A", "B", "C")))
''''     "A B C"
''''
''''   Unfolds outer array:
''''     >>> UnfoldParamArray(Array(Array("A", "B", "C")))
''''     Array("A", "B", "C")
''''     >>> ?Join(UnfoldParamArray(Array(Array("A", "B", "C"))))
''''     "A B C"
''''
'@Description "Unfolds a ParamArray argument when passed from another ParamArray."
Public Function UnfoldParamArray(ByVal ParamArrayArg As Variant) As Variant
Attribute UnfoldParamArray.VB_Description = "Unfolds a ParamArray argument when passed from another ParamArray."
    Guard.NotArray ParamArrayArg
    
    Dim DoUnfold As Boolean
    DoUnfold = (ArrayLib.NumberOfArrayDimensions(ParamArrayArg) = 1) And (LBound(ParamArrayArg) = 0) And (UBound(ParamArrayArg) = 0)
    If DoUnfold Then DoUnfold = IsArray(ParamArrayArg(0))
    If DoUnfold Then DoUnfold = ((ArrayLib.NumberOfArrayDimensions(ParamArrayArg(0)) = 1) And (LBound(ParamArrayArg(0), 1) = 0))
    If DoUnfold Then
        UnfoldParamArray = ParamArrayArg(0)
    Else
        UnfoldParamArray = ParamArrayArg
    End If
End Function


Public Function GetVarType(ByRef Variable As Variant) As String
    Dim NDim As String
    NDim = IIf(IsArray(Variable), "/Array", vbNullString)
    
    Dim TypeOfVar As VBA.VbVarType
    TypeOfVar = VarType(Variable) And Not vbArray

    Dim ScalarType As String
    Select Case TypeOfVar
        Case vbEmpty
            ScalarType = "vbEmpty"
        Case vbNull
            ScalarType = "vbNull"
        Case vbInteger
            ScalarType = "vbInteger"
        Case vbLong
            ScalarType = "vbLong"
        Case vbSingle
            ScalarType = "vbSingle"
        Case vbDouble
            ScalarType = "vbDouble"
        Case vbCurrency
            ScalarType = "vbCurrency"
        Case vbDate
            ScalarType = "vbDate"
        Case vbString
            ScalarType = "vbString"
        Case vbObject
            ScalarType = "vbObject"
        Case vbError
            ScalarType = "vbError"
        Case vbBoolean
            ScalarType = "vbBoolean"
        Case vbVariant
            ScalarType = "vbVariant"
        Case vbDataObject
            ScalarType = "vbDataObject"
        Case vbDecimal
            ScalarType = "vbDecimal"
        Case vbByte
            ScalarType = "vbByte"
        Case vbUserDefinedType
            ScalarType = "vbUserDefinedType"
        Case Else
            ScalarType = "vbUnknown"
    End Select
    GetVarType = ScalarType & NDim
End Function
