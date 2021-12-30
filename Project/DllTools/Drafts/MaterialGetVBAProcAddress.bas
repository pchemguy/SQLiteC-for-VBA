Attribute VB_Name = "MaterialGetVBAProcAddress"
'@Folder "DllTools.Drafts"
'@IgnoreModule
Option Explicit

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = "Library\" & LIB_NAME & "\Memtools\"

#If VBA7 Then
Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Private Declare PtrSafe Function GetRoutineAddress Lib "MemToolsLib" (ByVal ProcPtr As LongPtr) As LongPtr
#Else
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetRoutineAddress Lib "MemToolsLib" (ByVal ProcPtr As Long) As Long
#End If


Private Type TMemToolsGetVBAProcAddress
    DllMan As DllManager
End Type
Private this As TMemToolsGetVBAProcAddress


Private Sub RunPrintFuncAddress()
    Debug.Print CStr(GetFuncAddress(AddressOf GetMyAddressCopyMem))
End Sub

Public Sub GetMyAddressCopyMem()
    #If VBA7 Then
        Dim FuncPtr As LongPtr
    #Else
        Dim FuncPtr As Long
    #End If
    CopyMem FuncPtr, AddressOf GetMyAddressCopyMem, LenB(FuncPtr)
    Debug.Print CStr(FuncPtr)
End Sub

Private Sub GetMyAddressMemToolsLib()
    LoadMemToolsLib
    Debug.Print CStr(GetRoutineAddress(AddressOf GetMyAddressCopyMem))
End Sub

#If VBA7 Then
Private Function GetFuncAddress(ByVal FuncAddress As LongPtr) As LongPtr
#Else
Private Function GetFuncAddress(ByVal FuncAddress As Long) As Long
#End If
    GetFuncAddress = FuncAddress
End Function

Private Sub LoadMemToolsLib()
    Dim DllPath As String
    DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & ARCH
    Dim DllName As String
    DllName = "MemToolsLib.dll"
    Set this.DllMan = DllManager.Create(DllPath, DllName)
End Sub
