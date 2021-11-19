Attribute VB_Name = "DllCallPerformance"
'@Folder "DllManager.Demo.Custom and Extended DLL.DLL Call Performance"
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit
Option Private Module

''''
Private Const CYCLE_COUNT As Long = 10 ^ 7

#If Win64 Then
Private Declare PtrSafe Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare PtrSafe Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare PtrSafe Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare PtrSafe Function PerfGauge Lib "MemToolsLib" (ByVal OuterForCount As Long, ByVal InnerForCount As Long) As Long
#Else
Private Declare Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare Function PerfGauge Lib "MemToolsLib" (ByVal OuterForCount As Long, ByVal InnerForCount As Long) As Long
#End If

Private Type TDllCallPerformance
    DllMan As DllManager
End Type
Private this As TDllCallPerformance


Private Sub LoadDlls()
    Dim DllPath As String
    #If Win64 Then
        DllPath = "Library\" & ThisWorkbook.VBProject.Name & "\Demo - DLL - STDCALL and Adapter\memtools\x64"
    #Else
        DllPath = "Library\" & ThisWorkbook.VBProject.Name & "\Demo - DLL - STDCALL and Adapter\memtools\x32"
    #End If
    
    DllManager.Free
    DllManager.ForgetSingleton
    Dim DllName As String
    DllName = "MemToolsLib.dll"
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath, DllName, False)
    Set this.DllMan = DllMan
End Sub


Private Sub PerfGaugePerf()
    Const PROC_NAME As String = "PerfGaugePerf"
    LoadDlls
    
    Dim OuterCounter As Long
    Dim InnerCounter As Long
    Dim TimeDiffMs As Long
    
    OuterCounter = 10 ^ 5
    InnerCounter = 10 ^ 3
    TimeDiffMs = PerfGauge(OuterCounter, InnerCounter)
    
    Debug.Print PROC_NAME & ":" & " - " & Format$(OuterCounter, "#,##0") & " x " _
        & Format$(InnerCounter, "#,##0") & " times in " & TimeDiffMs & " ms"
End Sub


Private Sub DummySub0ArgsPerf()
    Const PROC_NAME As String = "DummySub0Args"
    LoadDlls
    
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    For CycleIndex = 0 To CYCLE_COUNT
        DummySub0Args
    Next CycleIndex
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    Debug.Print PROC_NAME & ":" & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub


Private Sub DummySub3ArgsPerf()
    Const PROC_NAME As String = "DummySub3Args"
    LoadDlls
    
    Dim Src() As Byte
    Dim Dst() As Byte
    Src = "ABCDEFGHIJKLMNOPGRSTUVWXYZ"
    Dst = String(255, "_")
    Dim SrcLen As String
    SrcLen = (UBound(Src) - LBound(Src) + 1 + Len(vbNullChar)) * 2
    
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    For CycleIndex = 0 To CYCLE_COUNT
        DummySub3Args Dst(0), Src(0), SrcLen
    Next CycleIndex
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    Debug.Print PROC_NAME & ":" & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub


Private Sub DummyFnc0ArgsPerf()
    Const PROC_NAME As String = "DummyFnc0Args"
    LoadDlls
    
    Dim Result As Long
    
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    For CycleIndex = 0 To CYCLE_COUNT
        Result = DummyFnc0Args
    Next CycleIndex
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    Debug.Print PROC_NAME & ":" & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub


Private Sub DummyFnc3ArgsPerf()
    Const PROC_NAME As String = "DummyFnc3Args"
    LoadDlls
    
    Dim Src() As Byte
    Dim Dst() As Byte
    Src = "ABCDEFGHIJKLMNOPGRSTUVWXYZ"
    Dst = String(255, "_")
    Dim SrcLen As String
    SrcLen = (UBound(Src) - LBound(Src) + 1 + Len(vbNullChar)) * 2
    Dim Result As Long
    
    Dim Start As Single
    Start = Timer
    Dim CycleIndex As Long
    For CycleIndex = 0 To CYCLE_COUNT
        Result = DummyFnc3Args(Dst(0), Src(0), SrcLen)
    Next CycleIndex
    Dim Delta As Long
    Delta = Round((Timer - Start) * 1000, 0)
    Debug.Print PROC_NAME & ":" & " - " & Format$(CYCLE_COUNT, "#,##0") _
        & " times in " & Delta & " ms"
End Sub
