Attribute VB_Name = "DllPerfRun"
'@Folder "DllTools.Performance"
Option Explicit


'@EntryPoint
'@Ignore UseMeaningfulName
Private Sub RunnerSub0()
    Dim DummyMax As Long
    DummyMax = 10 ^ 7
    
    Dim PerfTool As DllPerfLib
    Set PerfTool = DllPerfLib.Create(DummyMax)
    
    Dim TimeDiffMs As Long
    Dim LoopIndex As Long
    With PerfTool
        .TogglePrint
        
        Dim AverageCountDLL As Long
        AverageCountDLL = 2
    
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountDLL
            TimeDiffMs = TimeDiffMs + .Sub0ArgsDLLVBA(, TARGET_DLL)
        Next LoopIndex
        If AverageCountDLL > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountDLL
            Debug.Print "Sub0ArgsDLLVBA/DLL" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
    End With
End Sub


'@EntryPoint
Private Sub Runner()
    Dim GaugeMax As Long
    GaugeMax = 10 ^ 9
    Dim DummyMax As Long
    DummyMax = 10 ^ 7
    
    Dim PerfTool As DllPerfLib
    Set PerfTool = DllPerfLib.Create(DummyMax, GaugeMax)
    
    Dim TimeDiffMs As Long
    Dim LoopIndex As Long
    With PerfTool
        .TogglePrint
        
        Dim AverageCountGAU As Long
        Dim AverageCountDLL As Long
        Dim AverageCountVBA As Long
        AverageCountGAU = 20
        AverageCountDLL = 2
        AverageCountVBA = 10
        '''' ========== PerfGauge ========== ''''
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountGAU
            TimeDiffMs = TimeDiffMs + .PerfGaugeGet
        Next LoopIndex
        If AverageCountGAU > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountGAU
            Debug.Print "PerfGauge" & ":" & " - " & Format$(GaugeMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        '''' ---------- PerfGauge ---------- ''''
    
        '''' ========== Sub0ArgsDLLVBA ========== ''''
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountDLL
            TimeDiffMs = TimeDiffMs + .Sub0ArgsDLLVBA(, TARGET_DLL)
        Next LoopIndex
        If AverageCountDLL > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountDLL
            Debug.Print "Sub0ArgsDLLVBA/DLL" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountVBA
            TimeDiffMs = TimeDiffMs + .Sub0ArgsDLLVBA(, TARGET_VBA)
        Next LoopIndex
        If AverageCountVBA > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountVBA
            Debug.Print "Sub0ArgsDLLVBA/VBA" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        '''' ---------- Sub0ArgsDLLVBA ---------- ''''
    
        '''' ========== Sub3ArgsDLLVBA ========== ''''
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountDLL
            TimeDiffMs = TimeDiffMs + .Sub3ArgsDLLVBA(, TARGET_DLL)
        Next LoopIndex
        If AverageCountDLL > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountDLL
            Debug.Print "Sub3ArgsDLLVBA/DLL" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountVBA
            TimeDiffMs = TimeDiffMs + .Sub3ArgsDLLVBA(, TARGET_VBA)
        Next LoopIndex
        If AverageCountVBA > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountVBA
            Debug.Print "Sub3ArgsDLLVBA/VBA" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        '''' ---------- Sub3ArgsDLLVBA ---------- ''''
    
        '''' ========== Fnc0ArgsDLLVBA ========== ''''
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountDLL
            TimeDiffMs = TimeDiffMs + .Fnc0ArgsDLLVBA(, TARGET_DLL)
        Next LoopIndex
        If AverageCountDLL > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountDLL
            Debug.Print "Fnc0ArgsDLLVBA/DLL" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountVBA
            TimeDiffMs = TimeDiffMs + .Fnc0ArgsDLLVBA(, TARGET_VBA)
        Next LoopIndex
        If AverageCountVBA > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountVBA
            Debug.Print "Fnc0ArgsDLLVBA/VBA" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        '''' ---------- Fnc0ArgsDLLVBA ---------- ''''
    
        '''' ========== Fnc3ArgsDLLVBA ========== ''''
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountDLL
            TimeDiffMs = TimeDiffMs + .Fnc3ArgsDLLVBA(, TARGET_DLL)
        Next LoopIndex
        If AverageCountDLL > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountDLL
            Debug.Print "Fnc3ArgsDLLVBA/DLL" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        
        TimeDiffMs = 0
        For LoopIndex = 1 To AverageCountVBA
            TimeDiffMs = TimeDiffMs + .Fnc3ArgsDLLVBA(, TARGET_VBA)
        Next LoopIndex
        If AverageCountVBA > 0 Then
            TimeDiffMs = TimeDiffMs / AverageCountVBA
            Debug.Print "Fnc3ArgsDLLVBA/VBA" & ":" & " - " & Format$(DummyMax, "#,##0") & _
                " times in " & TimeDiffMs & " ms"
        End If
        DoEvents
        '''' ---------- Fnc3ArgsDLLVBA ---------- ''''
    End With
End Sub

