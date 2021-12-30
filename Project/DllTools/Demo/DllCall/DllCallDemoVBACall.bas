Attribute VB_Name = "DllCallDemoVBACall"
'@Folder "DllTools.Demo.DllCall"
''''
'''' WARNING: Dll calls can crash the application. With calls via DispCallFunc,
'''' the VBA compiler cannot perform any correctness checks on the target call.
'''' Make sure your work is saved and be prepared for Excel crashing.
''''
Option Explicit

Private Type TModuleState
    LongVal As Long
    LongRef As Long
    ByteVal As Byte
    ByteRef As Byte
    StrVal As String
    StrRef As String
End Type
'@Ignore MoveFieldCloserToUsage
Private this As TModuleState


'''' This demo calls a VBA function with the following signature:
''''   Private Function SixParamOneReturn(ByVal ByteVal As Byte, ByVal LongVal As Long, ByVal StrVal As String,
''''                                      ByRef ByteRef As Byte, ByRef LongRef As Long, ByRef StrRef As String) As Long
'''' While there is little sense in calling a VBA routine indirectly, this demo
'''' provides a convinient means for controlling the entire process, including
'''' transfer of parameters and returning values via the return variable and
'''' ByRef arguments. For illustrative and testing purposes, using module-level
'''' variables within the TModuleState structure.
'''' If successful, all printed lines should contain "MATCHED" mark.
''''
'''' Loosely follows https://akihitoyamashiro.blogspot.com/2020/07/how-to-use-function-pointer-in-vba-2.html
''''
'@EntryPoint
Private Sub Main()
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString)
    DllMan.CacheProcPtr "DllCallDemoVBACall", "In3Out3Ret1", AddressOf In3Out3Ret1
    
    With this
        .ByteVal = 10
        .LongVal = 30
        .StrVal = "StrVal"
        .ByteRef = 20
        .LongRef = 40
        .StrRef = "StrRef"
    End With
        
    With this
        Dim Arguments As Variant
        Arguments = Array( _
            .ByteVal, _
            .LongVal, _
            .StrVal, _
            VarPtr(.ByteRef), _
            VarPtr(.LongRef), _
            VarPtr(.StrRef) _
        )
    End With
    
    Debug.Print "==================== In3Out3Ret1 ===================="
    Dim Result As Long
    Result = DllMan.IndirectCall("DllCallDemoVBACall", "In3Out3Ret1", _
                                 CC_STDCALL, vbLong, Arguments)
    
    Debug.Print vbNewLine & "----- VERIFYING RETURNED VALUES -----"
    With this
        Debug.Print "ByteVal = " & CStr(.ByteVal) & vbTab & vbTab & IIf(.ByteVal = 10, "MATCHED/UNCHANGED", "MISMATCHED")
        Debug.Print "ByteRef = " & CStr(.ByteRef) & vbTab & vbTab & IIf(.ByteRef = 200, "MATCHED/UPDATED", "MISMATCHED")
        Debug.Print "LongVal = " & CStr(.LongVal) & vbTab & vbTab & IIf(.LongVal = 30, "MATCHED/UNCHANGED", "MISMATCHED")
        Debug.Print "LongRef = " & CStr(.LongRef) & vbTab & vbTab & IIf(.LongRef = 400, "MATCHED/UPDATED", "MISMATCHED")
        Debug.Print "StrVal  = " & CStr(.StrVal) & vbTab & IIf(.StrVal = "StrVal", "MATCHED/UNCHANGED", "MISMATCHED")
        Debug.Print "StrRef  = " & CStr(.StrRef) & vbTab & IIf(.StrRef = "StrRefNew", "MATCHED/UPDATED", "MISMATCHED")
    End With
    Debug.Print "Result  = " & CStr(Result) & vbTab & vbTab & IIf(Result = 70, "MATCHED", "MISMATCHED")
    Debug.Print "-------------------- In3Out3Ret1 --------------------"
End Sub


'@Ignore AssignedByValParameter, UseMeaningfulName
Private Function In3Out3Ret1(ByVal ByteVal As Byte, ByVal LongVal As Long, ByVal StrVal As String, _
                             ByRef ByteRef As Byte, ByRef LongRef As Long, ByRef StrRef As String) As Long
    Debug.Print "----- VERIFYING RECEIVED ARGUEMNTS -----"
    Debug.Print "ByteVal = " & CStr(ByteVal) & vbTab & vbTab & IIf(ByteVal = 10, "MATCHED", "MISMATCHED")
    Debug.Print "ByteRef = " & CStr(ByteRef) & vbTab & vbTab & IIf(ByteRef = 20, "MATCHED", "MISMATCHED")
    Debug.Print "LongVal = " & CStr(LongVal) & vbTab & vbTab & IIf(LongVal = 30, "MATCHED", "MISMATCHED")
    Debug.Print "LongRef = " & CStr(LongRef) & vbTab & vbTab & IIf(LongRef = 40, "MATCHED", "MISMATCHED")
    Debug.Print "StrVal  = " & CStr(StrVal) & vbTab & IIf(StrVal = "StrVal", "MATCHED", "MISMATCHED")
    Debug.Print "StrRef  = " & CStr(StrRef) & vbTab & IIf(StrRef = "StrRef", "MATCHED", "MISMATCHED")
    In3Out3Ret1 = LongVal + LongRef
    
    LongVal = 300
    LongRef = 400
    ByteVal = 100
    ByteRef = 200
    StrVal = "StrValNew"
    StrRef = "StrRefNew"
End Function
