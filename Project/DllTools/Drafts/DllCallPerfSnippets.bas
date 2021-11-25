Attribute VB_Name = "DllCallPerfSnippets"
'@Folder "DllTools.Drafts"
'@IgnoreModule
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "DllManager"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = _
    "Library" & PATH_SEP & LIB_NAME & PATH_SEP & _
    "Demo - DLL - STDCALL and Adapter" & PATH_SEP
Private Const CYCLE_COUNT As Long = 10 ^ 7

#If Win64 Then
Private Declare PtrSafe Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare PtrSafe Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare PtrSafe Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare PtrSafe Function PerfGauge Lib "MemToolsLib" (ByVal GaugeForCount As Long) As Long
#Else
Private Declare Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare Function PerfGauge Lib "MemToolsLib" (ByVal GaugeForCount As Long) As Long
#End If

Private Type TDllCallPerformance
    DllMan As DllManager
End Type
Private this As TDllCallPerformance


Private Sub LoadDlls()
    Dim DllPath As String
    #If Win64 Then
        DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & "memtools\x64"
    #Else
        DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & "memtools\x32"
    #End If
    
    DllManager.Free
    Dim DllName As String
    DllName = "MemToolsLib.dll"
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath, DllName)
    Set this.DllMan = DllMan
End Sub


Private Sub PerfGaugePerf()
    Const PROC_NAME As String = "PerfGaugePerf"
    LoadDlls
    
    Dim GaugeCounter As Long
    Dim TimeDiffMs As Long
    
    GaugeCounter = 10 ^ 8
    TimeDiffMs = PerfGauge(GaugeCounter)
    
    Debug.Print PROC_NAME & ":" & " - " & Format$(GaugeCounter, "#,##0") & _
        " times in " & TimeDiffMs & " ms"
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

