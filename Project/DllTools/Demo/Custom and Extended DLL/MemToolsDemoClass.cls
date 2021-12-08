VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemToolsDemoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "DllTools.Demo.Custom and Extended DLL"
Option Explicit

Private SomeArray(0 To 10 ^ 4) As Long
#If VBA7 Then
    Private SomeArrayBase As LongPtr
#Else
    Private SomeArrayBase As Long
#End If
Private SomeArrayElementSize As Long


Private Sub Class_Initialize()
    SomeArray(1023) = -1
    SomeArrayBase = VarPtr(SomeArray(0))
    SomeArrayElementSize = Len(SomeArray(0))
End Sub


#If VBA7 Then
Public Property Get SomeArrayAddress() As LongPtr
#Else
Public Property Get SomeArrayAddress() As Long
#End If
    SomeArrayAddress = SomeArrayBase
End Property


#If VBA7 Then
Public Function GetMemLong(ByVal LongAddress As LongPtr) As Long
#Else
Public Function GetMemLong(ByVal LongAddress As Long) As Long
#End If
    GetMemLong = SomeArray(CLng(LongAddress - SomeArrayBase) \ SomeArrayElementSize)
End Function


#If VBA7 Then
'@Ignore ProcedureCanBeWrittenAsFunction: Function is above
Public Sub GetMemLongSub(ByVal LongAddress As LongPtr, ByRef Dest As Long)
#Else
'@Ignore ProcedureCanBeWrittenAsFunction: Function is above
Public Sub GetMemLongSub(ByVal LongAddress As Long, ByRef Dest As Long)
#End If
    Dest = SomeArray(CLng(LongAddress - SomeArrayBase) \ SomeArrayElementSize)
End Sub
