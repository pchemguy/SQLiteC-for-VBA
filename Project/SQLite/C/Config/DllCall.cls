VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DllCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Call dll routine via dll/routine names."
'@Folder "SQLite.C.Config"
'@ModuleDescription "Call dll routine via dll/routine names."
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
''''
'''' WARNING: Dll calls can crash the application. With calls via DispCallFunc,
'''' the VBA compiler cannot perform any correctness checks on the target call.
'''' Make sure your work is saved and be prepared for Excel crashing.
''''
Option Explicit

Private Const MODULE_NAME As String = "DllCall"

#If VBA7 Then
Private Declare PtrSafe Function DispCallFunc Lib "OleAut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As CALLCONV, _
    ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As LongPtr, ByVal prgpvarg As LongPtr, ByVal pvargResult As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
#Else
Private Declare Function DispCallFunc Lib "OleAut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal CC As CALLCONV, _
    ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
#End If

Private Type TObjectState
    DllMan As DllManager
    ProcPtrCache As Scripting.Dictionary
End Type
Private this As TObjectState


'@DefaultMember
Public Function Create(ByVal DllMan As DllManager) As DllCall
Attribute Create.VB_UserMemId = 0
    Dim Instance As DllCall
    Set Instance = New DllCall
    Instance.Init DllMan
    Set Create = Instance
End Function


Friend Sub Init(ByVal DllMan As DllManager)
    Guard.NullReference DllMan
    Set this.DllMan = DllMan
    Set this.ProcPtrCache = New Scripting.Dictionary
    this.ProcPtrCache.CompareMode = TextCompare
End Sub


#If VBA7 Then
Friend Sub CacheProcPtr(ByVal ModuleName As String, ByVal ProcName As String, ByVal ProcPtr As LongPtr)
#Else
Friend Sub CacheProcPtr(ByVal ModuleName As String, ByVal ProcName As String, ByVal ProcPtr As Long)
#End If
    Dim FileName As String
    FileName = LCase$(ModuleName) & IIf(ModuleName Like "*.???", vbNullString, ".dll")
    this.ProcPtrCache(FileName & "." & LCase$(ProcName)) = ProcPtr
End Sub


#If VBA7 Then
Friend Function ProcAddressGet(ByVal DllName As String, ByVal ProcName As String) As LongPtr
    Dim DllPtr As LongPtr
    Dim ProcPtr As LongPtr
#Else
Friend Function ProcAddressGet(ByVal DllName As String, ByVal ProcName As String) As Long
    Dim DllPtr As Long
    Dim ProcPtr As Long
#End If
    Dim ProcKey As String
    Dim FileName As String
    If DllName Like "*.???" Then
        FileName = LCase$(DllName)
    Else
        FileName = LCase$(DllName) & ".dll"
    End If
    ProcKey = LCase$(FileName & "." & ProcName)
    With this.ProcPtrCache
        If .Exists(ProcKey) Then
            ProcAddressGet = .Item(ProcKey)
            Exit Function
        End If
    End With
    
    With this.DllMan.Dlls
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


Public Function IndirectCall( _
            ByVal ModuleName As String, _
            ByVal ProcName As Variant, _
            ByVal CallnigConv As CALLCONV, _
            ByVal ReturnType As VbVarType, _
            ByRef Arguments As Variant) As Variant
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
