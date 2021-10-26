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
'@Folder "DllManager"
'@ModuleDescription "Loads/unloads dll libraries"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule SelfAssignedDeclaration: ok for FileSystemObject
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
    Private Declare PtrSafe Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" _
        (ByVal hModule As LongPtr, ByVal lpFilename As LongPtr, ByVal nSize As Long) As Long
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
    Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" _
        (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
#End If

Private Const VERBOUS As Boolean = False

Public Enum DllLoadStatus
    LOAD_OK = -1
    LOAD_FAIL = 0
    LOAD_ALREADY_LOADED = 1
End Enum

Private Const ERROR_BAD_EXE_FORMAT As Long = 193
Private Const LoadingDllErr As Long = 48

Private Type TDllManager
    DefaultPath As String
    Dlls As Scripting.Dictionary '''' Dll file name to handle map
End Type
Private this As TDllManager


Public Function Create(Optional ByVal DefaultPath As String = vbNullString, _
                       Optional ByVal DllNames As Variant = Empty) As DllManager
    Dim Instance As DllManager
    Set Instance = New DllManager
    Instance.Init DefaultPath, DllNames
    Set Create = Instance
End Function


Friend Sub Init(Optional ByVal DefaultPath As String = vbNullString, _
                Optional ByVal DllNames As Variant = Empty)
    Set this.Dlls = New Scripting.Dictionary
    this.Dlls.CompareMode = TextCompare
    Me.DefaultPath = DefaultPath
    If Not IsEmpty(DllNames) Then
        If Not LoadMultiple(DllNames) = LOAD_OK Then
            Err.Raise LoadingDllErr, "DllManager/Constructor", "Failed to load Dlls"
        End If
    End If
End Sub


Private Sub Class_Terminate()
    ResetDllSearchPath
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
    If fso.FolderExists(Value) Then
        '''' Absolute existing path is provided
        Path = Value
    ElseIf fso.FolderExists(fso.BuildPath(ThisWorkbook.Path, Value)) Then
        '''' Relative existing path is provided
        Path = fso.BuildPath(ThisWorkbook.Path, Value)
    Else
        Err.Raise ErrNo.FileNotFoundErr, "DllManager", _
                  "DefaultPath not found: <" & Value & ">"
    End If
    Path = fso.GetAbsolutePathName(Path)
    
    '''' Set the default dll directory for LoadLibrary
    '''' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setdlldirectorya#remarks
    Dim ExecStatus As Boolean
    ExecStatus = SetDllDirectory(Path)
    If ExecStatus Then
        this.DefaultPath = Path
    Else
        Debug.Print "SetDllDirectory failed. Error code: " & CStr(Err.LastDllError)
    End If
End Property


'''' https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setdlldirectorya#remarks
Public Sub ResetDllSearchPath()
    Dim ExecStatus As Boolean
    ExecStatus = SetDllDirectory(vbNullString)
    If ExecStatus Then
        this.DefaultPath = vbNullString
    Else
        Debug.Print "Reset SetDllDirectory failed. Error code: " & CStr(Err.LastDllError)
    End If
End Sub


Public Function Load(ByVal DllName As String, _
            Optional ByVal Path As String = vbNullString, _
            Optional ByVal ResolvePath As Boolean = True) As DllLoadStatus
    Guard.EmptyString DllName
    
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = fso.GetFileName(DllName)
    
    If this.Dlls.Exists(FileName) Then
        Debug.Print "A library with file name <" & FileName & "> has already been loaded."
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
                Err.Raise ErrNo.FileNotFoundErr, "DllManager", _
                          "Path not found: <" & Path & ">"
            End If
        End If
        Prefix = fso.GetAbsolutePathName(Prefix)
        
        If fso.FileExists(DllName) Then
            FilePathName = DllName
        ElseIf fso.FileExists(fso.BuildPath(Prefix, DllName)) Then
            FilePathName = fso.BuildPath(Prefix, DllName)
        ElseIf fso.FileExists(fso.BuildPath(this.DefaultPath, DllName)) Then
            FilePathName = fso.BuildPath(this.DefaultPath, DllName)
        Else
            Err.Raise ErrNo.FileNotFoundErr, "DllManager", _
                      "DllName not found: <" & DllName & ">"
        End If
        FilePathName = fso.GetAbsolutePathName(FilePathName)
    Else
        FilePathName = DllName
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
        If VERBOUS Then Debug.Print "<" & DllName & "> loaded."
        Load = LOAD_OK
    Else
        LastDllError = Err.LastDllError
        Debug.Print "Library <" & FilePathName & "> loading error: " & CStr(LastDllError)
        If LastDllError = ERROR_BAD_EXE_FORMAT Then
            '''' TODO: Make PE Format analyzer class to check target architechture
            ''''       and imports (dependencies).
            Err.Raise LoadingDllErr, "DllManager", "LoadLibrary error 193: <" _
                      & FilePathName & ">" & vbNewLine & _
                      "Check that DLL bitness matches that of Office."
        End If
        Load = LOAD_FAIL
    End If
End Function


Public Function Free(Optional ByVal DllName As String) As DllLoadStatus
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = fso.GetFileName(DllName)
    Dim Result As Long
    If Not this.Dlls.Exists(FileName) Then
        Debug.Print "<" & DllName & "> not loaded."
        Free = LOAD_OK
    Else
        Result = FreeLibrary(this.Dlls(FileName))
        If Result <> 0 Then
            If VERBOUS Then Debug.Print "<" & DllName & "> unloaded."
            Free = LOAD_OK
            this.Dlls.Remove FileName
        Else
            Free = LOAD_FAIL
            Debug.Print "Error unloading <" & DllName & ">. Result: " _
                & CStr(Result) & ". LastDllError: "; CStr(Err.LastDllError)
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
    If -Abs(Result) = LOAD_OK Then Debug.Print "DLLs loaded"
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
    If Result = LOAD_OK Then Debug.Print "DLLs unloaded"
    FreeMultiple = Result
End Function


Public Function GetDllPath(ByVal DllName As String) As String
    Guard.EmptyString DllName
    If Not this.Dlls.Exists(DllName) Then
        GetDllPath = vbNullString
        Exit Function
    End If
    Const MAX_PATH As Long = 260
    Dim PathName As String
    PathName = Space$(MAX_PATH + 1)
    Dim PathLength As Long
    PathLength = GetModuleFileName(this.Dlls(DllName), StrPtr(PathName), MAX_PATH)
    If PathLength = 0 Then Err.Raise ErrNo.UnknownClassErr, "DllManager", "Failed to obtain Dll path."
    GetDllPath = PathName
End Function
