Attribute VB_Name = "FuncPtr"
'@Folder "DllTools.Drafts"
'@IgnoreModule
Option Explicit

'''' https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comtypes.callconv

#If VBA7 Then
Private Declare PtrSafe Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As CALLCONV, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByVal prgvt As LongPtr, _
    ByVal prgpvarg As LongPtr, _
    ByVal pvargResult As LongPtr _
    ) As Long
#Else
Private Declare Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal CC As CALLCONV, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByVal prgvt As Long, _
    ByVal prgpvarg As Long, _
    ByVal pvargResult As Long _
    ) As Long
#End If

Private ViaDispCall As Boolean

Private Sub CalledFunction()
    Stop
    'MsgBox "DispCallFunc just called me!"
End Sub


Private Sub CalledBlankFunction()
End Sub


Private Sub RunCalledFunction()
    Dim DispCallFuncResult As Long
    Dim Result As Variant
    Result = vbEmpty
    DispCallFuncResult = DispCallFunc( _
        0, _
        AddressOf CalledFunction, _
        CLng(4), _
        VbVarType.vbEmpty, _
        0, _
        0, _
        0, _
        VarPtr(Result))

    Dim DummyMax As Long
    DummyMax = 10 ^ 5
    ViaDispCall = True
        
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    If ViaDispCall Then
        For CycleIndex = 0 To DummyMax
            DispCallFuncResult = DispCallFunc( _
                0, _
                AddressOf CalledBlankFunction, _
                CLng(4), _
                VbVarType.vbEmpty, _
                0, _
                0, _
                0, _
                VarPtr(Result))
        Next CycleIndex
    Else
        For CycleIndex = 0 To DummyMax
            CalledBlankFunction
        Next CycleIndex
    End If
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    Debug.Print IIf(ViaDispCall, "DispCallFunc", "Direct") & ":" & " - " & Format$(DummyMax, "#,##0") _
        & " times in " & TimeDiffMs & " ms"
End Sub
