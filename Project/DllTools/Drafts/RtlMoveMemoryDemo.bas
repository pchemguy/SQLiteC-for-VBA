Attribute VB_Name = "RtlMoveMemoryDemo"
'@Folder "DllTools.Drafts"
'@IgnoreModule
Option Explicit

Private Const CRYPT_STRING_BINARY As Long = &H2&

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
    Private Declare PtrSafe Function ToString Lib "Crypt32" Alias "CryptBinaryToStringA" ( _
        ByRef Source As Any, ByVal NumBytes As Long, ByVal Flags As Long, ByRef Destination As Any, ByRef BytesWritten As Long) As Long
#Else
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As Long)
#End If


Private Sub CopyMemDemo()
    Dim Src As Long
    Dim Dst As Long
    Src = 10241024
    Dst = 0
    Dim ByteCount As Long
    ByteCount = LenB(Src)
    Dim DummyMax As Long
    DummyMax = 10 ^ 7
    
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    For CycleIndex = 0 To DummyMax
        CopyMem Dst, Src, ByteCount
    Next CycleIndex
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    Debug.Print "CopyMem" & ":" & " - " & Format$(DummyMax, "#,##0") _
        & " times in " & TimeDiffMs & " ms"

COPY_MEM_DEMOO:
End Sub


#If VBA7 Then
Private Sub ToStringDemo()
    Dim Src As Long
    Dim Dst As Long
    Src = 10241024
    Dst = 0
    Dim ByteCount As Long
    ByteCount = LenB(Src)
    Dim BytesWritten As Long
    BytesWritten = LenB(Dst)
    Dim Result As Long
    
    Dim DummyMax As Long
    DummyMax = 10 ^ 6
    
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    For CycleIndex = 0 To DummyMax
        Result = ToString(Src, ByteCount, CRYPT_STRING_BINARY, Dst, BytesWritten)
    Next CycleIndex
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    Debug.Print "ToString" & ":" & " - " & Format$(DummyMax, "#,##0") _
        & " times in " & TimeDiffMs & " ms"
End Sub
#End If

