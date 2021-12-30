VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DllManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Loads/unloads dll libraries"
'@Folder "DllTools.Manager"
'@ModuleDescription "Loads/unloads dll libraries"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration: ok for FileSystemObject
Option Explicit

Private Const MODULE_NAME As String = "DllManager"

#If Win64 Then
Private Const ARCH As String = "x64"
#Else
Private Const ARCH As String = "x32"
#End If

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" _
        (ByVal hModule As LongPtr, ByVal lpFilename As LongPtr, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function DispCallFunc Lib "OleAut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As CALLCONV, _
        ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As LongPtr, ByVal prgpvarg As LongPtr, ByVal pvargResult As LongPtr) As Long
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" _
        (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
    Private Declare Function DispCallFunc Lib "OleAut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal CC As CALLCONV, _
        ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long
#End If

Private Type TObjectState
    DefaultPath As String
    '''' DllFileName => DllHandle map. Always convert filename to lower case and
    '''' include extension. If no extension is provided, add ".dll" as default.
    Dlls As Scripting.Dictionary
    ProcPtrCache As Scripting.Dictionary
End Type
Private this As TObjectState


Public Function Create(Optional ByVal DefaultPath As String = vbNullString, _
                       Optional ByVal DllNames As Variant = Empty, _
                       Optional ByVal ResolvePath As Boolean = True) As DllManager
    Dim Instance As DllManager
    Set Instance = New DllManager
    Instance.Init DefaultPath, DllNames, ResolvePath
    Set Create = Instance
End Function


Friend Sub Init(Optional ByVal DefaultPath As String = vbNullString, _
                Optional ByVal DllNames As Variant = Empty, _
                Optional ByVal ResolvePath As Boolean = True)
    Set this.Dlls = New Scripting.Dictionary
    this.Dlls.CompareMode = TextCompare
    Me.DefaultPath = IIf(Len(DefaultPath) > 0, DefaultPath, _
        "Library\" & ThisWorkbook.VBProject.Name & "\dll\" & ARCH)
    Select Case VarType(DllNames)
        Case vbString
            If Len(DllNames) > 0 Then _
                If Load(DllNames, , ResolvePath) <> LOAD_OK Then GoTo LOAD_FAIL:
        Case Is >= vbArray
            If LoadMultiple(DllNames) <> LOAD_OK Then GoTo LOAD_FAIL:
        Case Else
            If Not (IsEmpty(DllNames) Or IsNull(DllNames)) Then Debug.Print _
                "Unusable second argument is ignored."
    End Select
    Set this.ProcPtrCache = New Scripting.Dictionary
    this.ProcPtrCache.CompareMode = TextCompare
    
    Exit Sub
LOAD_FAIL:
    Err.Raise LoadingDllErr, MODULE_NAME & "/Constructor", "Failed to load Dlls"
End Sub


Private Sub Class_Terminate()
    Me.DefaultPath = vbNullString
    Dim ResultCode As DllLoadStatus
    ResultCode = FreeMultiple
    Debug.Assert ResultCode = LOAD_OK
End Sub


Public Property Get Dlls() As Scripting.Dictionary
    Set Dlls = this.Dlls
End Property


Public Property Get DefaultPath() As String
    DefaultPath = this.DefaultPath
End Property


Public Property Let DefaultPath(ByVal Value As String)
    Dim fso As New Scripting.FileSystemObject
    Dim Path As String
    If Len(Value) = 0 Then
        Path = vbNullString
    ElseIf fso.FolderExists(Value) Then
        '''' Absolute existing path is provided
        Path = fso.GetAbsolutePathName(Value)
    ElseIf fso.FolderExists(ThisWorkbook.Path & "\" & Value) Then
        '''' Relative existing path is provided
        Path = fso.GetAbsolutePathName(ThisWorkbook.Path & "\" & Value)
    Else
        Err.Raise ErrNo.FileNotFoundErr, MODULE_NAME, _
                  "DefaultPath not found: <" & Value & ">"
    End If
    
    '''' Set the default dll directory for LoadLibrary
    '''' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setdlldirectorya#remarks
    Dim ExecStatus As Boolean
    ExecStatus = SetDllDirectory(Path)
    If ExecStatus Then
        this.DefaultPath = Path
    Else
        Dim LogMsg As String
        LogMsg = "SetDllDirectory failed. Error code: " & CStr(Err.LastDllError)
        Logger.Logg LogMsg, , DEBUGLEVEL_ERROR
    End If
End Property


Public Function Load(ByVal DllName As String, _
            Optional ByVal Path As String = vbNullString, _
            Optional ByVal ResolvePath As Boolean = True) As DllLoadStatus
    Guard.EmptyString DllName
    
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = LCase$(fso.GetFileName(DllName))
    If Not FileName Like "*.???" Then FileName = FileName & ".dll"
    
    Dim LogMsg As String
    If this.Dlls.Exists(FileName) Then
        LogMsg = "A library with file name <" & FileName & _
                 "> has already been loaded."
        Logger.Logg LogMsg, , DEBUGLEVEL_WARN
        Load = LOAD_ALREADY_LOADED
        Exit Function
    End If
    
    Dim FilePathName As String
    If ResolvePath Then
        Dim Prefix As String
        If Len(Path) = 0 Then
            '''' Default path
            Prefix = this.DefaultPath
        Else
            If fso.FolderExists(Path) Then
                '''' Absolute existing path is provided
                Prefix = Path
            ElseIf fso.FolderExists(fso.BuildPath(ThisWorkbook.Path, Path)) Then
                '''' Relative existing path is provided
                Prefix = fso.BuildPath(ThisWorkbook.Path, Path)
            Else
                Err.Raise ErrNo.FileNotFoundErr, MODULE_NAME, _
                          "Path not found: <" & Path & ">"
            End If
        End If
        FilePathName = fso.BuildPath(Prefix, DllName)
        If fso.FileExists(FilePathName) Then
            FilePathName = fso.GetAbsolutePathName(FilePathName)
        Else
            Err.Raise ErrNo.FileNotFoundErr, MODULE_NAME, _
                "DllName not found: <" & DllName & ">"
        End If
    Else
        FilePathName = IIf(Len(Path) > 0, Path & "\", Path) & DllName
    End If
    
    Dim LastDllError As Long
    #If VBA7 Then
        Dim DllHandle As LongPtr
    #Else
        Dim DllHandle As Long
    #End If
    DllHandle = LoadLibrary(FilePathName)
    If DllHandle <> 0 Then
        this.Dlls(FileName) = DllHandle
        LogMsg = "<" & DllName & "> loaded."
        Logger.Logg LogMsg, , DEBUGLEVEL_INFO
        Load = LOAD_OK
    Else
        LastDllError = Err.LastDllError
        LogMsg = "Library <" & FilePathName & "> loading error: " & _
                 CStr(LastDllError)
        Logger.Logg LogMsg, , DEBUGLEVEL_ERROR
        If LastDllError = ERROR_BAD_EXE_FORMAT Then
            '''' TODO: Make PE Format analyzer class to check target
            ''''       architechture and imports (dependencies).
            Err.Raise LoadingDllErr, MODULE_NAME & "/Load", "LoadLibrary " & _
                "error " & ERROR_BAD_EXE_FORMAT & ": <" & FilePathName & ">" _
                & vbNewLine & "Check that DLL bitness matches that of Office."
        End If
        Load = LOAD_FAIL
    End If
End Function


Public Function Free(Optional ByVal DllName As String) As DllLoadStatus
    If this.Dlls Is Nothing Then Exit Function
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = LCase$(fso.GetFileName(DllName))
    If Not FileName Like "*.???" Then FileName = FileName & ".dll"
    Dim Result As Long
    Dim LogMsg As String
    If Not this.Dlls.Exists(FileName) Then
        LogMsg = "<" & DllName & "> not loaded."
        Logger.Logg LogMsg, , DEBUGLEVEL_WARN
        Free = LOAD_OK
    Else
        Result = FreeLibrary(this.Dlls(FileName))
        If Result <> 0 Then
            LogMsg = "<" & DllName & "> unloaded."
            Logger.Logg LogMsg, , DEBUGLEVEL_INFO
            Free = LOAD_OK
            this.Dlls.Remove FileName
        Else
            Free = LOAD_FAIL
            LogMsg = "Error unloading <" & DllName & ">. Result: " _
                & CStr(Result) & ". LastDllError: " & CStr(Err.LastDllError)
            Logger.Logg LogMsg, , DEBUGLEVEL_ERROR
        End If
    End If
End Function


Public Function LoadMultiple(ParamArray DllNames() As Variant) As DllLoadStatus
    Dim FileNames() As Variant
    FileNames = UnfoldParamArray(DllNames)
    Dim Result As Long
    Result = LOAD_OK
    Dim FileNameIndex As Long
    For FileNameIndex = LBound(FileNames) To UBound(FileNames)
        Result = Result And Load(FileNames(FileNameIndex))
    Next FileNameIndex
    If -Abs(Result) = LOAD_OK Then
        Dim LogMsg As String
        LogMsg = "DLLs loaded"
        Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    End If
    LoadMultiple = -Abs(Result)
End Function


'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function FreeMultiple(ParamArray DllNames() As Variant) As DllLoadStatus
    Dim FileNames() As Variant
    FileNames = UnfoldParamArray(DllNames)
    Dim Result As Long
    Result = LOAD_OK
    If UBound(FileNames) - LBound(FileNames) + 1 = 0 Then FileNames = this.Dlls.Keys
    Dim FileNameIndex As Long
    For FileNameIndex = LBound(FileNames) To UBound(FileNames)
        Result = Result And Free(FileNames(FileNameIndex))
    Next FileNameIndex
    If Result = LOAD_OK Then
        Dim LogMsg As String
        LogMsg = "DLLs unloaded"
        Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    End If
    FreeMultiple = Result
End Function


Public Function GetDllPath(ByVal DllName As String) As String
    Guard.EmptyString DllName
    GetDllPath = vbNullString
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = LCase$(fso.GetFileName(DllName))
    If Not FileName Like "*.???" Then FileName = FileName & ".dll"
    If this.Dlls Is Nothing Then Exit Function
    If Not this.Dlls.Exists(FileName) Then Exit Function
    Const MAX_PATH As Long = 260
    Dim PathName As String
    PathName = Space$(MAX_PATH + 1)
    Dim PathLength As Long
    PathLength = GetModuleFileName(this.Dlls(FileName), StrPtr(PathName), MAX_PATH)
    If PathLength = 0 Then Err.Raise ErrNo.UnknownClassErr, MODULE_NAME, _
                           "Failed to obtain Dll path."
    GetDllPath = PathName
End Function


#If VBA7 Then
Public Sub CacheProcPtr(ByVal ModuleName As String, ByVal ProcName As String, ByVal ProcPtr As LongPtr)
#Else
Public Sub CacheProcPtr(ByVal ModuleName As String, ByVal ProcName As String, ByVal ProcPtr As Long)
#End If
    Dim FileName As String
    FileName = LCase$(ModuleName) & IIf(ModuleName Like "*.???", vbNullString, ".dll")
    this.ProcPtrCache(FileName & "." & LCase$(ProcName)) = ProcPtr
End Sub


#If VBA7 Then
Public Function ProcAddressGet(ByVal DllName As String, ByVal ProcName As String) As LongPtr
    Dim DllPtr As LongPtr
    Dim ProcPtr As LongPtr
#Else
Public Function ProcAddressGet(ByVal DllName As String, ByVal ProcName As String) As Long
    Dim DllPtr As Long
    Dim ProcPtr As Long
#End If
    Dim ProcKey As String
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = LCase$(fso.GetFileName(DllName))
    If Not FileName Like "*.???" Then FileName = FileName & ".dll"
    ProcKey = LCase$(FileName & "." & ProcName)
    With this.ProcPtrCache
        If .Exists(ProcKey) Then
            ProcAddressGet = .Item(ProcKey)
            Exit Function
        End If
    End With
    
    With this.Dlls
        If .Exists(FileName) Then
            DllPtr = .Item(FileName)
        Else
            Err.Raise ErrNo.MemberNotExistErr, MODULE_NAME & "/ ProcAddressGet", _
                "DLL <" & DllName & "> has not been loaded or been unloaded."
        End If
    End With
    ProcPtr = GetProcAddress(DllPtr, ProcName)
    If ProcPtr = 0 Then
        Err.Raise ErrNo.MemberNotExistErr, MODULE_NAME & "/ ProcAddressGet", _
            "Could not obtain address of " & DllName & "." & ProcName & "."
    End If
    this.ProcPtrCache(ProcKey) = ProcPtr
    
    ProcAddressGet = ProcPtr
End Function


''''
'''' WARNING: Dll calls can crash the application. With calls via DispCallFunc,
'''' the VBA compiler cannot perform any correctness checks on the target call.
'''' Make sure your work is saved and be prepared for Excel crashing.
''''
'@Description "Calls a dll routine using dll/proc names via DispCallFunc."
Public Function IndirectCall( _
            ByVal ModuleName As String, _
            ByVal ProcName As Variant, _
            ByVal CallnigConv As CALLCONV, _
            ByVal ReturnType As VbVarType, _
            ByRef Arguments As Variant) As Variant
Attribute IndirectCall.VB_Description = "Calls a dll routine using dll/proc names via DispCallFunc."
    #If VBA7 Then
        Dim COMInstancePtr As LongPtr
        Dim ProcPtr As LongPtr
        Dim ReturnPtr As LongPtr
        Dim ArgPtrs() As LongPtr
        Dim ArgPtrsPtr As LongPtr
        Dim ArgTypesPtr As LongPtr
    #Else
        Dim COMInstancePtr As Long
        Dim ProcPtr As Long
        Dim ReturnPtr As Long
        Dim ArgPtrs() As Long
        Dim ArgPtrsPtr As Long
        Dim ArgTypesPtr As Long
    #End If
    
    COMInstancePtr = 0
    ProcPtr = ProcAddressGet(ModuleName, ProcName)
    
    Dim ReturnValue As Variant: ReturnValue = 0 '''' RB workaround
    ReturnPtr = IIf(ReturnType <> vbEmpty, VarPtr(ReturnValue), 0)
    
    '@Ignore IntegerDataType
    Dim ArgTypes() As Integer
    Dim ArgCount As Long
    Dim ArgIndex As Long
    If IsEmpty(Arguments) Then
        ArgCount = 0
        ArgPtrsPtr = 0
        ArgTypesPtr = 0
    Else
        Dim ArgsLB As Long
        Dim ArgsUB As Long
        ArgsLB = LBound(Arguments)
        ArgsUB = UBound(Arguments)
        ArgCount = ArgsUB - ArgsLB + 1
        ReDim ArgPtrs(ArgsLB To ArgsUB)
        ReDim ArgTypes(ArgsLB To ArgsUB)
        For ArgIndex = ArgsLB To ArgsUB
            ArgPtrs(ArgIndex) = VarPtr(Arguments(ArgIndex))
            ArgTypes(ArgIndex) = VarType(Arguments(ArgIndex))
        Next ArgIndex
        ArgPtrsPtr = VarPtr(ArgPtrs(ArgsLB))
        ArgTypesPtr = VarPtr(ArgTypes(ArgsLB))
    End If
    
    Dim CallResut As HRESULT
    CallResut = DispCallFunc(COMInstancePtr, ProcPtr, CallnigConv, ReturnType, _
                             ArgCount, ArgTypesPtr, ArgPtrsPtr, ReturnPtr)
    If CallResut <> S_OK Then
        Err.Raise CallResut, MODULE_NAME & "/IndirectCall", _
                  "HRESULT error 0x" & Hex$(CallResut)
    End If
    IndirectCall = ReturnValue
End Function
