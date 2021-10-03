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
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryW" (ByVal lpPathName As String) As Boolean
#End If

Private Const LOAD_OK As Long = -1
Private Const LOAD_FAIL As Long = 0
Private Const LOAD_ALREADY_LOADED As Long = 1

Private Type TDllManager
    DefaultPath As String
    Dlls As Scripting.Dictionary
End Type
Private this As TDllManager


'@DefaultMember
Public Function Create(Optional ByVal DefaultPath As String = vbNullString) As DllManager
Attribute Create.VB_UserMemId = 0
    Dim Instance As DllManager
    Set Instance = New DllManager
    Instance.Init DefaultPath
    Set Create = Instance
End Function


Friend Sub Init(Optional ByVal DefaultPath As String = vbNullString)
    Set this.Dlls = New Scripting.Dictionary
    this.Dlls.CompareMode = TextCompare
    Me.DefaultPath = DefaultPath
End Sub


Private Sub Class_Terminate()
    ResetDllSearchPath
    FreeMultiple
End Sub


Public Property Get Dlls() As Scripting.Dictionary
    Set Dlls = this.Dlls
End Property


Public Property Get DefaultPath() As String
    DefaultPath = this.DefaultPath
End Property


Public Property Let DefaultPath(ByVal Value As String)
    '@Ignore SelfAssignedDeclaration: Stateless utility object
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


Public Function Load(ByVal DllName As String, Optional ByVal Path As String = vbNullString) As Long
    '@Ignore SelfAssignedDeclaration: Stateless utility object
    Dim fso As New Scripting.FileSystemObject
    Dim FileName As String
    FileName = fso.GetFileName(DllName)
    
    If this.Dlls.Exists(FileName) Then
        Debug.Print "A library with file name <" & FileName & "> has already been loaded."
        Load = LOAD_ALREADY_LOADED
        Exit Function
    End If
    
    Dim Prefix As String
    If fso.FolderExists(Path) Then
        '''' Absolute existing path is provided
        Prefix = Path
    ElseIf fso.FolderExists(fso.BuildPath(ThisWorkbook.Path, Path)) Then
        '''' Relative existing path is provided
        Prefix = fso.BuildPath(ThisWorkbook.Path, Path)
    Else
        '''' Default path
        Prefix = fso.BuildPath(fso.BuildPath(ThisWorkbook.Path, "Library"), _
                               ThisWorkbook.VBProject.Name)
    End If
    Prefix = fso.GetAbsolutePathName(Prefix)
    
    
    Dim FilePathName As String
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
    
    Dim DllHandle As Long
    DllHandle = LoadLibrary(FilePathName)
    If DllHandle <> 0 Then
        '@Ignore IndexedDefaultMemberAccess: Dictionary
        this.Dlls(FileName) = DllHandle
        Debug.Print "<" & DllName & "> loaded."
        Load = LOAD_OK
    Else
        Debug.Print "Library <" & FilePathName & "> loading error: " & CStr(Err.LastDllError)
        Load = LOAD_FAIL
    End If
End Function


Public Function Free(Optional ByVal DllName As String) As Long
    '@Ignore SelfAssignedDeclaration: Stateless utility object
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
            Debug.Print "<" & DllName & "> unloaded."
            Free = LOAD_OK
            this.Dlls.Remove FileName
        Else
            Free = LOAD_FAIL
            Debug.Print "Error unloading <" & DllName & ">. Result: " _
                & CStr(Result) & ". LastDllError: "; CStr(Err.LastDllError)
        End If
    End If
End Function


Public Function LoadMultiple(ParamArray DllNames() As Variant) As Long
    Dim Result As Long
    Result = LOAD_OK
    Dim FileNames() As Variant
    FileNames = UnfoldParamArray(DllNames)
    Dim FileNameIndex As Long
    For FileNameIndex = LBound(FileNames) To UBound(FileNames)
        Result = Result And Load(FileNames(FileNameIndex))
    Next FileNameIndex
    LoadMultiple = -Abs(Result)
End Function


Public Function FreeMultiple(ParamArray DllNames() As Variant) As Long
    Dim Result As Long
    Result = LOAD_OK
    Dim FileNames() As Variant
    FileNames = UnfoldParamArray(DllNames)
    If UBound(FileNames) - LBound(FileNames) + 1 = 0 Then FileNames = this.Dlls.Keys
    Dim FileNameIndex As Long
    For FileNameIndex = LBound(FileNames) To UBound(FileNames)
        Result = Result And Free(FileNames(FileNameIndex))
    Next FileNameIndex
    FreeMultiple = Result
End Function
