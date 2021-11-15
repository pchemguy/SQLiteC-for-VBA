Attribute VB_Name = "DemoLibMemory"
'@Folder "SQLiteDBdev.MemoryTools"
'@IgnoreModule: This is part of the https://github.com/cristianbuse/VBA-MemoryTools
Option Explicit
Option Private Module

Private Const LOOPS As Long = 10 ^ 7

Private Sub DemoMain()
    DemoMem
    Debug.Print String(55, "-")
    DemoMemByteSpeed
    DemoMemIntSpeed
    DemoMemLongSpeed
    DemoMemLongPtrSpeed
    DemoMemObjectSpeed
End Sub

Private Sub DemoInstanceRedirection()
    Const loopsCount As Long = 100000
    Dim i As Long
    Dim T As Double
    '
    T = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory2(i).ID = i
    Next i
    Debug.Print "Public  Init (seconds): " & VBA.Round(Timer - T, 3)
    '
    T = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory(i).ID = i
    Next i
    Debug.Print "Private Init (seconds): " & VBA.Round(Timer - T, 3)
End Sub

Private Sub DemoMem()
    #If VBA7 Then
        Dim ptr As LongPtr
    #Else
        Dim ptr As Long
    #End If
    Dim i As Long
    Dim Arr() As Variant
    ptr = ObjPtr(Application)
    '
    'Read Memory using MemByte
    ReDim Arr(0 To PTR_SIZE - 1)
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = MemByte(ptr + i)
    Next i
    Debug.Print Join(Arr, " ")
    '
    'Read Memory using MemInt
    ReDim Arr(0 To PTR_SIZE / 2 - 1)
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = MemInt(ptr + i * 2)
    Next i
    Debug.Print Join(Arr, " ")
    '
    'Read Memory using MemLong
    ReDim Arr(0 To PTR_SIZE / 4 - 1)
    For i = LBound(Arr) To UBound(Arr)
        Arr(i) = MemLong(ptr + i * 4)
    Next i
    Debug.Print Join(Arr, " ")
    '
    'Read Memory using MemLongPtr
    Debug.Print MemLongPtr(ptr)
    '
    'Write Memory using MemByte
    ptr = 0
    MemByte(VarPtr(ptr)) = 24
    Debug.Assert ptr = 24
    MemByte(VarPtr(ptr) + 2) = 24
    Debug.Assert ptr = 1572888
    '
    'Write Memory using MemInt
    ptr = 0
    MemInt(VarPtr(ptr) + 2) = 300
    Debug.Assert ptr = 19660800
    '
    'Write Memory using MemLong
    ptr = 0
    MemLong(VarPtr(ptr)) = 77777
    Debug.Assert ptr = 77777
    '
    'Write Memory using MemLongPtr
    MemLongPtr(VarPtr(ptr)) = ObjPtr(Application)
    Debug.Assert ptr = ObjPtr(Application)
End Sub

Private Sub DemoMemByteSpeed()
    Dim x1 As Byte: x1 = 1
    Dim x2 As Byte: x2 = 2
    Dim i As Long
    Dim T As Double
    '
    T = Timer
    For i = 1 To LOOPS
        MemByte(VarPtr(x1)) = MemByte(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    Dim ByteCount As Long
    ByteCount = Len(x1)
    T = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, ByteCount
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemIntSpeed()
    Dim x1 As Integer: x1 = 11111
    Dim x2 As Integer: x2 = 22222
    Dim i As Long
    Dim T As Double
    '
    T = Timer
    For i = 1 To LOOPS
        MemInt(VarPtr(x1)) = MemInt(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    Dim ByteCount As Long
    ByteCount = Len(x1)
    T = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, ByteCount
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemLongSpeed()
    Dim x1 As Long: x1 = 111111111
    Dim x2 As Long: x2 = 222222222
    Dim i As Long
    Dim T As Double
    '
    T = Timer
    For i = 1 To LOOPS
        MemLong(VarPtr(x1)) = MemLong(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    Dim ByteCount As Long
    ByteCount = Len(x1)
    T = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, ByteCount
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemLongPtrSpeed()
    #If WIN64 Then
        Dim x1 As LongLong: x1 = 111111111111111^
        Dim x2 As LongLong: x2 = 111111111111112^
    #Else
        Dim x1 As Long: x1 = 111111111
        Dim x2 As Long: x2 = 222222222
    #End If
    Dim i As Long
    Dim T As Double
    '
    T = Timer
    For i = 1 To LOOPS
        MemLongPtr(VarPtr(x1)) = MemLongPtr(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    Dim ByteCount As Long
    ByteCount = Len(x1)
    T = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, ByteCount
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemObjectSpeed()
    Dim i As Long
    Dim T As Double
    Dim D As DemoClass: Set D = New DemoClass
    Dim Obj As Object
    #If WIN64 Then
        Dim ptr As LongLong
    #Else
        Dim ptr As Long
    #End If
    '
    ptr = ObjPtr(D)
    T = Timer
    For i = 1 To LOOPS
        Set Obj = MemObject(ptr)
    Next i
    Debug.Print "Dereferenced an Object " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - T, 3) & " seconds"
    DoEvents
End Sub
