Attribute VB_Name = "MemToolsDemo"
'@Folder "DllTools.Demo.Custom and Extended DLL"
'@IgnoreModule: It follows https://github.com/cristianbuse/VBA-MemoryTools
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = "Library\" & LIB_NAME & "\Memtools\"

''''Copy <Long> By native assign   10,000,000 times in 51 ms
''''Copy <Long> From array native  10,000,000 times in 121 ms
''''Copy <Long> From array by sub  10,000,000 times in 641 ms
''''Copy <Long> From class by sub  10,000,000 times in 879 ms
''''Copy <Long> From class by func 10,000,000 times in 1012 ms
''''Copy <Long> By API             10,000,000 times in 0.34 seconds
''''Copy <Long> By Ref             10,000,000 times in 3.914 seconds
''''Copy <Long> By CopyMem         10,000,000 times in 102 ms
''''
''''Copy <Long> By native assign   10,000,000 times in 55 ms
''''Copy <Long> From array native  10,000,000 times in 137 ms
''''Copy <Long> From array by sub  10,000,000 times in 645 ms
''''Copy <Long> From class by sub  10,000,000 times in 1062 ms
''''Copy <Long> From class by func 10,000,000 times in 1113 ms
''''Copy <Long> By API             10,000,000 times in 20.719 seconds
''''Copy <Long> By Ref             10,000,000 times in 3.867 seconds

#If Win64 Then
Private Declare PtrSafe Sub CopyMem Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
Private Declare Sub CopyMem Lib "MemToolsLib" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Private Const LOOPS As Long = 10 ^ 7

Private SomeArrayM(0 To 10 ^ 4) As Long
    #If VBA7 Then
        Private SomeArrayBaseM As LongPtr
    #Else
        Private SomeArrayBaseM As Long
    #End If
Private SomeArrayElementSizeM As Long

Private Type TMemToolsLibDemo
    DllMan As DllManager
End Type
Private this As TMemToolsLibDemo


Private Sub TestCopyLong()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    #If Win64 Then
        DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & "x64"
    #Else
        DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & "x32"
    #End If
    LoadDlls DllPath
    
    Dim x1 As Long
    x1 = 111111111
    Dim x2 As Long
    x2 = 222222222
    Dim ByteCount As Long
    ByteCount = Len(x1)
    Dim T As Single
    Dim i As Long
    T = Timer
    For i = 1 To LOOPS
        CopyMem x1, x2, ByteCount
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By CopyMem " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"
    T = Timer
    For i = 1 To LOOPS
        x1 = x2
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By native  " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"
    
    Set this.DllMan = Nothing
End Sub


Private Sub TestCopyLongFromClass()
    Dim x1 As Long
    Dim x2 As Long
    Dim x3 As Long
    Dim T As Single
    Dim i As Long
    
    Dim SomeArray(0 To 10 ^ 4) As Long
    #If VBA7 Then
        Dim SomeArrayBase As LongPtr
        Dim TargetAddress As LongPtr
    #Else
        Dim SomeArrayBase As Long
        Dim TargetAddress As Long
    #End If
    Dim SomeArrayElementSize As Long

    Dim Instance As MemToolsDemoClass
    Set Instance = New MemToolsDemoClass
    
    ''''=======================================================================
    
    T = Timer
    For i = 1 To LOOPS
        x1 = x2
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By native assign   " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"
    
    ''''=======================================================================
    
    SomeArray(1023) = -1
    SomeArrayBase = VarPtr(SomeArray(0))
    SomeArrayElementSize = Len(SomeArray(0))
    TargetAddress = SomeArrayBase + (1024 - 1) * 4
 
    x2 = SomeArray(CLng(TargetAddress - SomeArrayBase) \ SomeArrayElementSize)
    T = Timer
    For i = 1 To LOOPS
        x2 = SomeArray(CLng(TargetAddress - SomeArrayBase) \ SomeArrayElementSize)
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> From array native  " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"
    
    ''''=======================================================================
    
    SomeArrayM(1023) = -1
    SomeArrayBaseM = VarPtr(SomeArrayM(0))
    SomeArrayElementSizeM = Len(SomeArrayM(0))
    TargetAddress = SomeArrayBaseM + (1024 - 1) * 4
    
    T = Timer
    For i = 1 To LOOPS
        CopyLong TargetAddress, x3
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> From array by sub  " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"

    ''''=======================================================================
    
    TargetAddress = Instance.SomeArrayAddress + (1024 - 1) * 4
    
    T = Timer
    For i = 1 To LOOPS
        Instance.GetMemLongSub TargetAddress, x1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> From class by sub  " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"

    ''''=======================================================================
    
    TargetAddress = Instance.SomeArrayAddress + (1024 - 1) * 4
    
    T = Timer
    For i = 1 To LOOPS
        x1 = Instance.GetMemLong(TargetAddress)
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> From class by func " & Format$(LOOPS, "#,##0") _
        & " times in " & Round((Timer - T) * 1000, 0) & " ms"
    
    ''''=======================================================================
End Sub


#If VBA7 Then
Public Sub CopyLong(ByVal LongAddress As LongPtr, ByRef Dest As Long)
#Else
Public Sub CopyLong(ByVal LongAddress As Long, ByRef Dest As Long)
#End If
    Dest = SomeArrayM(CLng(LongAddress - SomeArrayBaseM) \ SomeArrayElementSizeM)
End Sub


Private Sub LoadDlls(ByVal DllPath As String)
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
    Set this.DllMan = DllMan
    Dim DllName As Variant
    DllName = "MemToolsLib.dll"
    DllMan.Load DllName
End Sub
