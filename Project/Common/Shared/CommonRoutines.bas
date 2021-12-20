Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
'@IgnoreModule MoveFieldCloserToUsage, IndexedDefaultMemberAccess, ProcedureNotUsed
Option Explicit

#If Win64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
    Public Const vbLongLong As Long = 20&
#End If

Private lastID As Double


Public Function GetTimeStampMs() As String
    '''' On Windows, the Timer resolution is subsecond, the fractional part (the four characters at the end
    '''' given the format) is concatenated with DateTime. It appears that the Windows' high precision time
    '''' source available via API yields garbage for the fractional part.
    GetTimeStampMs = Format$(Now, "yyyy-MM-dd HH:mm:ss") & Right$(Format$(Timer, "#0.000"), 4)
End Function


'''' The number of seconds since the Epoch is multiplied by 10^4 to bring the first
'''' four fractional places in Timer value into the whole part before truncation.
'''' Long does not provide sufficient number of digits, so returning double.
'''' Alternatively, a Currency type could be used.
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

Public Function GetEpoch() As Double
    GetEpoch = CDbl(DateDiff("s", DateSerial(1970, 1, 1), Date)) + Timer
End Function

Public Function EpochToString(Optional ByVal Epoch As Double = -1) As String
    Dim EpochValue As Double
    Dim EpochRef As Date
    EpochRef = DateSerial(1970, 1, 1)
    If Epoch > 0 Then
        EpochValue = Epoch
    Else
        EpochValue = CDbl(DateDiff("s", EpochRef, Date)) + Timer
    End If
    Dim DateVal As Date
    DateVal = DateAdd("s", EpochValue, EpochRef)
    EpochToString = Format$(DateVal, "YYYY-MM-DD hh:mm:ss") & "." & _
                    Format$(Round((EpochValue - Int(EpochValue)) * 1000), "000")
End Function

Public Function GenerateGUID() As String
    GenerateGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function

Public Function RandomLong() As Long
    RandomLong = Val("&H" & Left$(GenerateGUID, 8))
End Function


'''' When sub/function captures a list of arguments in a ParamArray and passes it
'''' to the next routine expecting a list of arguments, the second routine receives
'''' a 2D array instead of 1D with the outer dimension having a single element.
'''' This function check the arguments and unfolds the outer dimesion as necessary.
'''' Any function accepting a ParamArray argument should be able to use it.
''''
'''' Unfold if the following conditions are satisfied:
''''     - ParamArrayArg is a 1D array
''''     - UBound(ParamArrayArg, 1) = LBound(ParamArrayArg, 1) = 0
''''     - ParamArrayArg(0) is a 1D 0-based array
''''
'''' Return
''''     - ParamArrayArg(0), if unfolding is necessary
''''     - ParamArrayArg, if ParamArrayArg is array, but not all conditions are satisfied
'''' Raise an error if is not an array
'@Description "Unfolds a ParamArray argument when passed from another ParamArray."
Public Function UnfoldParamArray(ByVal ParamArrayArg As Variant) As Variant
Attribute UnfoldParamArray.VB_Description = "Unfolds a ParamArray argument when passed from another ParamArray."
    Guard.NotArray ParamArrayArg
    Dim DoUnfold As Boolean
    DoUnfold = (ArrayLib.NumberOfArrayDimensions(ParamArrayArg) = 1) And _
               (LBound(ParamArrayArg) = 0) And (UBound(ParamArrayArg) = 0)
    If DoUnfold Then DoUnfold = IsArray(ParamArrayArg(0))
    If DoUnfold Then
        DoUnfold = ((ArrayLib.NumberOfArrayDimensions(ParamArrayArg(0)) = 1) And _
                   (LBound(ParamArrayArg(0), 1) = 0))
    End If
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


'''' Resolves file pathname
''''
'''' This helper routines attempts to interpret provided pathname as
'''' a file reference:
'''' 1) check if provided reference is a valid absolute file pathname, if not,
'''' 2) if AllowNonExistent and parent folder exists and the reference does not
''''      point to an existing folder, return FilePathName, possibly prefixed
''''      with ThisWorkbook.Path
'''' If absolute path is provided and not resolved, fail resolution at this point.
'''' 3) construct an array of possible file locations:
''''      - ThisWorkbook.Path & Application.PathSeparator
''''      - Environ("APPDATA") & Application.PathSeparator &
''''          & ThisWorkbook.VBProject.Name & Application.PathSeparator
''''    construct an array of possible file names:
''''      - FilePathName
''''          skip if len=0, or prefix is not relative
''''      - FilePathName & Ext (Ext comes from the second argument)
'''' 4) loop through all possible path/filename combinations until a valid
''''    pathname is found or all options are exhausted
''''
'''' Args:
''''   FilePathName (string):
''''     File pathname
''''
''''   DefaultExts (string or string/array):
''''     1D array of default extensions or a single default extension
''''
''''   AllowNonExistent (boolean, optional, False):
''''     If set to True, FilePathName may point to a non-existent file.
''''     Check with ThisWorkbook.Path prefix
''''     If FilePathName is blank, raise an error
''''
'''' Returns:
''''   String:
''''     Resolved valid absolute pathname pointing to an existing file.
''''
'''' Raises:
''''   Err.FileNotFoundErr:
''''     If provided pathname cannot be resolved to a valid file pathname.
''''
'''' Examples:
''''   >>> ?VerifyOrGetDefaultPath(Environ$("ComSpec"), "")
''''   "C:\Windows\system32\cmd.exe"
''''
''''   Raises error:
''''     >>> ?VerifyOrGetDefaultPath("")
''''     Raises "FileNotFoundErr" error
''''
''''     TODO: Add unit tests
''''     >>> ?VerifyOrGetDefaultPath("___.___")
''''     Raises "FileNotFoundErr" error
''''
''''     TODO: Add unit tests
''''     >>> ?VerifyOrGetDefaultPath("___.___", Array("___"), True)
''''     Raises "FileNotFoundErr" error
''''
''''     >>> ?VerifyOrGetDefaultPath("", , True)
''''     Raises "FileNotFoundErr" error
''''
''''   TODO: Add unit tests
''''   Allow non-existent:
''''     >>> ?VerifyOrGetDefaultPath("___.___", , True)
''''     "<Thisworkbook.Path>\___.___"
''''
''''     >>> ?VerifyOrGetDefaultPath(Application.PathSeparator & "___.___", , True)
''''     Application.PathSeparator & "___.___"
''''
'@Description "Resolves file pathname"
Public Function VerifyOrGetDefaultPath( _
                    ByVal FilePathName As String, _
           Optional ByVal DefaultExts As Variant = Empty, _
           Optional ByVal AllowNonExistent As Boolean = False) As String
Attribute VerifyOrGetDefaultPath.VB_Description = "Resolves file pathname"
    Dim PATHuSEP As String
    PATHuSEP = Application.PathSeparator
    Dim PROJuNAME As String
    PROJuNAME = ThisWorkbook.VBProject.Name
    
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    Dim PathNameCandidate As String
    
    '''' === (1) === Check if FilePathName is a valid path to an existing file.
    If fso.FileExists(FilePathName) Then
        VerifyOrGetDefaultPath = fso.GetAbsolutePathName(FilePathName)
        Exit Function
    End If
    
    '''' === (2) === Return FilePathName pointing to a non-existing file.
    If AllowNonExistent Then
        If Len(FilePathName) = 0 Then GoTo FILE_NOT_FOUND
        If Len(fso.GetDriveName(FilePathName)) = 0 Then
            '''' Blank drive name - relative path
            PathNameCandidate = fso.GetAbsolutePathName( _
                                    ThisWorkbook.Path & PATHuSEP & FilePathName)
        Else
            PathNameCandidate = fso.GetAbsolutePathName(FilePathName)
        End If
        If Not fso.FolderExists(fso.GetParentFolderName(PathNameCandidate)) Then
            GoTo FILE_NOT_FOUND
        End If
        VerifyOrGetDefaultPath = PathNameCandidate
        Exit Function
    End If
    
    '''' Absolute path name, if valid, should be resolved in (1) or (2)
    If Len(fso.GetDriveName(FilePathName)) > 0 Then GoTo FILE_NOT_FOUND
    
    '''' === (3a) === Array of prefixes
    Dim Prefixes As Variant
    Prefixes = Array( _
        ThisWorkbook.Path & PATHuSEP & "Library" & PATHuSEP & PROJuNAME & PATHuSEP, _
        ThisWorkbook.Path & PATHuSEP, _
        Environ$("APPDATA") & PATHuSEP & PROJuNAME & PATHuSEP _
    )
    
    '''' === (3b) === Array of filenames
    Dim NameCount As Long
    NameCount = 0
    
    Dim UseFilePathName As Boolean
    UseFilePathName = Len(FilePathName) > 1
    If UseFilePathName Then
        NameCount = NameCount + 1
    End If
    If VarType(DefaultExts) = vbString Then
        If Len(DefaultExts) > 0 Then NameCount = NameCount + 1
    ElseIf VarType(DefaultExts) >= vbArray Then
        NameCount = NameCount + UBound(DefaultExts, 1) - LBound(DefaultExts, 1) + 1
        Debug.Assert VarType(DefaultExts(0)) = vbString
    End If
    
    If NameCount = 0 Then GoTo FILE_NOT_FOUND
    
    Dim FileNames() As String
    ReDim FileNames(0 To NameCount - 1)
    Dim ExtIndex As Long
    Dim FileNameIndex As Long
    FileNameIndex = 0
    If UseFilePathName Then
        FileNames(FileNameIndex) = FilePathName
        FileNameIndex = FileNameIndex + 1
    End If
    If VarType(DefaultExts) = vbString Then
        If Len(DefaultExts) > 0 Then
            FileNames(FileNameIndex) = FilePathName & "." & DefaultExts
        End If
    ElseIf VarType(DefaultExts) >= vbArray Then
        For ExtIndex = LBound(DefaultExts, 1) To UBound(DefaultExts, 1)
            FileNames(FileNameIndex) = FilePathName & "." & DefaultExts(ExtIndex)
            FileNameIndex = FileNameIndex + 1
        Next ExtIndex
    End If
    
    '''' === (4) === Loop through pathnames
    Dim PrefixIndex As Long
    
    For FileNameIndex = 0 To UBound(FileNames)
        For PrefixIndex = 0 To UBound(Prefixes)
            PathNameCandidate = Prefixes(PrefixIndex) & FileNames(FileNameIndex)
            If fso.FileExists(PathNameCandidate) Then
                VerifyOrGetDefaultPath = fso.GetAbsolutePathName(PathNameCandidate)
                Exit Function
            End If
        Next PrefixIndex
    Next FileNameIndex
    '''' If reached this point, proceed to FILE_NOT_FOUND
    
FILE_NOT_FOUND:
    Err.Raise ErrNo.FileNotFoundErr, "CommonRoutines", _
              "File <" & FilePathName & "> not found!"
End Function


'''' Tests if argument is falsy
''''
'''' Falsy values:
''''   Numeric: 0
''''   String:  vbNullString
''''   Variant: Empty
''''   Object:  Nothing
''''   Boolean: False
''''   Null:    Null
''''
'''' Args:
''''   arg:
''''     Value to be tested for falsiness
''''
'''' Returns:
''''   True, if "arg" is Falsy
''''   Flase, if "arg" is Truthy (not Falsy)
''''
'''' Examples:
''''   >>> ?IsFalsy(0.0#)
''''   True
''''
''''   >>> ?IsFalsy(0.1)
''''   False
''''
''''   >>> ?IsFalsy(Null)
''''   True
''''
''''   >>> ?IsFalsy(Empty)
''''   True
''''
''''   >>> ?IsFalsy(False)
''''   True
''''
''''   >>> ?IsFalsy(Nothing)
''''   True
''''
''''   >>> ?IsFalsy("")
''''   True
''''
'@Description("Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing")
Public Function IsFalsy(ByVal arg As Variant) As Boolean
Attribute IsFalsy.VB_Description = "Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing"
    Select Case VarType(arg)
        Case vbEmpty, vbNull
            IsFalsy = True
        Case vbInteger, vbLong, vbSingle, vbDouble
            IsFalsy = Not CBool(arg)
        Case vbString
            IsFalsy = (arg = vbNullString)
        Case vbObject
            IsFalsy = (arg Is Nothing)
        Case vbBoolean
            IsFalsy = Not arg
        Case Else
            IsFalsy = False
    End Select
End Function


'''' Places a 2D array with top header on a worksheet
''''
'''' The top row is set in bold and centered, columns are adjusted with autofit.
''''
'''' Args:
''''   DataArray (2D array, variant):
''''     Data to be placed on a worksheet
''''
''''   TopLeftCell (Excel.Range):
''''     Reference to the top left corner of the target area.
''''
''''   NoTopHeader (boolean):
''''     If true, do not format the first row as header
''''
'@Description "Places a 2D array with top header on a worksheet"
Public Sub Array2Range(ByVal DataArray As Variant, _
                       ByVal TopLeftCell As Excel.Range, _
              Optional ByVal NoTopHeader As Boolean = False)
Attribute Array2Range.VB_Description = "Places a 2D array with top header on a worksheet"
    Guard.NotArray DataArray
    Guard.NullReference TopLeftCell
    Guard.ExpressionErr ArrayLib.NumberOfArrayDimensions(DataArray) = 2, _
                        ExpectedArrayErr, "CommonRoutines", "Expected 2D Array"
    
    Dim OutRange As Excel.Range
    Set OutRange = TopLeftCell.Resize( _
        UBound(DataArray, 1) - LBound(DataArray, 1) + 1, _
        UBound(DataArray, 2) - LBound(DataArray, 2) + 1 _
    )
    With OutRange
        .Value = DataArray
        If Not NoTopHeader Then
            .Rows(1).HorizontalAlignment = xlCenter
            .Rows(1).Font.Bold = True
        End If
        .Columns.AutoFit
    End With
End Sub
