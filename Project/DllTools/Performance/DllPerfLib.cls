VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DllPerfLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides utilities for guaging performance of DLL calls via the memtools DLL."
'@Folder "DllTools.Performance"
'@ModuleDescription "Provides utilities for guaging performance of DLL calls via the memtools DLL."
'@PredeclaredId
'@IgnoreModule ParameterCanBeByVal: Testing ByRef
'@IgnoreModule VariableNotUsed, ParameterNotUsed, EmptyMethod
Option Explicit

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = "Library\" & LIB_NAME & "\Memtools\"

Public Enum TargetTypeEnum
    TARGET_DLL = 0&
    TARGET_VBA = 1&
End Enum

#If Win64 Then
Private Declare PtrSafe Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare PtrSafe Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare PtrSafe Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare PtrSafe Function PerfGauge Lib "MemToolsLib" (ByVal ForCount As Long) As Long
#Else
Private Declare Sub DummySub0Args Lib "MemToolsLib" ()
Private Declare Sub DummySub3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DummyFnc0Args Lib "MemToolsLib" () As Long
Private Declare Function DummyFnc3Args Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare Function PerfGauge Lib "MemToolsLib" (ByVal ForCount As Long) As Long
#End If

Private Type TDllPerfLib
    DllMan As DllManager
    DummyForCount As Long
    GaugeForCount As Long
    PrintToImmediate As Boolean
    Src() As Byte
    Dst() As Byte
    SrcLen As Long
End Type
Private this As TDllPerfLib


Public Function Create( _
            Optional ByVal DummyForCount As Long = 10000000, _
            Optional ByVal GaugeForCount As Long = 10000000) As DllPerfLib
    Dim Instance As DllPerfLib
    Set Instance = New DllPerfLib
    Instance.Init DummyForCount, GaugeForCount
    Set Create = Instance
End Function

Friend Sub Init(Optional ByVal DummyForCount As Long = 10000000, _
                Optional ByVal GaugeForCount As Long = 10000000)
    With this
        .DummyForCount = DummyForCount
        .GaugeForCount = GaugeForCount
        .PrintToImmediate = True
    End With
    LoadDlls

    this.Src = "ABCDEFGHIJKLMNOPGRSTUVWXYZ"
    this.Dst = String(255, "_")
    this.SrcLen = (UBound(this.Src) - LBound(this.Src) + 1 + Len(vbNullChar)) * 2
End Sub

Private Sub Class_Terminate()
    UnLoadDlls
End Sub

Public Sub TogglePrint()
    this.PrintToImmediate = Not this.PrintToImmediate
End Sub

Private Sub LoadDlls()
    Dim DllPath As String
    DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & ARCH
    Dim DllName As String
    DllName = "MemToolsLib.dll"
    Set this.DllMan = DllManager.Create(DllPath, DllName)
End Sub

Private Sub UnLoadDlls()
    '@Ignore FunctionReturnValueDiscarded
    this.DllMan.FreeMultiple
    Set this.DllMan = Nothing
End Sub

Public Function PerfGaugeGet(Optional ByVal GaugeForCount As Long = -1) As Long
    Dim GaugeMax As Long
    GaugeMax = IIf(GaugeForCount > 0, GaugeForCount, this.GaugeForCount)
    Dim TimeDiffMs As Long
    TimeDiffMs = PerfGauge(GaugeMax)
    If this.PrintToImmediate Then
        Debug.Print "PerfGauge" & ":" & " - " & Format$(GaugeMax, "#,##0") & _
            " times in " & TimeDiffMs & " ms"
    End If
    PerfGaugeGet = TimeDiffMs
End Function

Public Function Sub0ArgsDLLVBA(Optional ByVal DummyForCount As Long = -1, _
            Optional ByVal TargetType As TargetTypeEnum = TARGET_DLL) As Long
    Dim DummyMax As Long
    DummyMax = IIf(DummyForCount > 0, DummyForCount, this.DummyForCount)
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    If TargetType = TARGET_DLL Then
        For CycleIndex = 0 To DummyMax
            DummySub0Args
        Next CycleIndex
    Else
        For CycleIndex = 0 To DummyMax
            DummySub0ArgsVBA
        Next CycleIndex
    End If
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    Dim Source As String
    Source = "Sub0ArgsDLLVBA/" & Array("DLL", "VBA")(TargetType)
    If this.PrintToImmediate Then
        Debug.Print Source & ":" & " - " & Format$(DummyMax, "#,##0") _
            & " times in " & TimeDiffMs & " ms"
    End If
    Sub0ArgsDLLVBA = TimeDiffMs
End Function

Public Function Sub3ArgsDLLVBA(Optional ByVal DummyForCount As Long = -1, _
            Optional ByVal TargetType As TargetTypeEnum = TARGET_DLL) As Long
    Dim Src() As Byte
    Src = this.Src
    Dim Dst() As Byte
    Dst = this.Dst
    Dim SrcLen As Long
    SrcLen = this.SrcLen
    Dim DummyMax As Long
    DummyMax = IIf(DummyForCount > 0, DummyForCount, this.DummyForCount)
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    If TargetType = TARGET_DLL Then
        For CycleIndex = 0 To DummyMax
            DummySub3Args Dst(0), Src(0), SrcLen
        Next CycleIndex
    Else
        For CycleIndex = 0 To DummyMax
            DummySub3ArgsVBA Dst(0), Src(0), SrcLen
        Next CycleIndex
    End If
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    If this.PrintToImmediate Then
        Debug.Print "Sub3ArgsDLLVBA" & ":" & " - " & Format$(DummyMax, "#,##0") _
            & " times in " & TimeDiffMs & " ms"
    End If
    Sub3ArgsDLLVBA = TimeDiffMs
End Function

Public Function Fnc0ArgsDLLVBA(Optional ByVal DummyForCount As Long = -1, _
            Optional ByVal TargetType As TargetTypeEnum = TARGET_DLL) As Long
    Dim Result As Long
    Dim DummyMax As Long
    DummyMax = IIf(DummyForCount > 0, DummyForCount, this.DummyForCount)
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    If TargetType = TARGET_DLL Then
        For CycleIndex = 0 To DummyMax
            Result = DummyFnc0Args
        Next CycleIndex
    Else
        For CycleIndex = 0 To DummyMax
            Result = DummyFnc0ArgsVBA
        Next CycleIndex
    End If
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    If this.PrintToImmediate Then
        Debug.Print "Fnc0ArgsDLLVBA" & ":" & " - " & Format$(DummyMax, "#,##0") _
            & " times in " & TimeDiffMs & " ms"
    End If
    Fnc0ArgsDLLVBA = TimeDiffMs
End Function

Public Function Fnc3ArgsDLLVBA(Optional ByVal DummyForCount As Long = -1, _
            Optional ByVal TargetType As TargetTypeEnum = TARGET_DLL) As Long
    Dim Src() As Byte
    Src = this.Src
    Dim Dst() As Byte
    Dst = this.Dst
    Dim SrcLen As Long
    SrcLen = this.SrcLen
    Dim Result As Long
    Dim DummyMax As Long
    DummyMax = IIf(DummyForCount > 0, DummyForCount, this.DummyForCount)
    Dim CycleIndex As Long
    Dim Start As Single
    Start = Timer
    If TargetType = TARGET_DLL Then
        For CycleIndex = 0 To DummyMax
            Result = DummyFnc3Args(Dst(0), Src(0), SrcLen)
        Next CycleIndex
    Else
        For CycleIndex = 0 To DummyMax
            Result = DummyFnc3ArgsVBA(Dst(0), Src(0), SrcLen)
        Next CycleIndex
    End If
    Dim TimeDiffMs As Long
    TimeDiffMs = Round((Timer - Start) * 1000, 0)
    If this.PrintToImmediate Then
        Debug.Print "Fnc3ArgsDLLVBA" & ":" & " - " & Format$(DummyMax, "#,##0") _
            & " times in " & TimeDiffMs & " ms"
    End If
    Fnc3ArgsDLLVBA = TimeDiffMs
End Function

'''' ===== PERFORMANCE TEST STUBS ===== ''''

Private Sub DummySub0ArgsVBA()
End Sub

Private Sub DummySub3ArgsVBA(ByRef Destination As Byte, _
            ByRef Source As Byte, ByVal Length As Long)
End Sub

Private Function DummyFnc0ArgsVBA() As Long
    Dim Result As Long
    Result = 10241024
    DummyFnc0ArgsVBA = Result
End Function

Private Function DummyFnc3ArgsVBA(ByRef Destination As Byte, _
            ByRef Source As Byte, ByVal Length As Long) As Long
    Dim Result As Long
    Result = 10241024
    DummyFnc3ArgsVBA = Result
End Function

