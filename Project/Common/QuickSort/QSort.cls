VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides QuickSort routines."
'@IgnoreModule FunctionReturnValueDiscarded
'@Folder "Common.QuickSort"
'@ModuleDescription "Provides QuickSort routines."
'@PredeclaredId
Option Explicit
Option Compare Text


'
' Sorts an array of numbers or strings using quicksort algorithm.
' It first creates an array of indices and orders it.
' Once sorting is complete, the ordered array of indices
' is used to order the original array. (This indirect
' approach is used to reduce the number of swaps of the
' elements in the original array, which can be strings
' and, in general, objects.)
' Ascii strings are sorted in case insensitive fashion (Option Compare Text).
'
Public Function Vector(ByRef ValueArray As Variant, _
              Optional ByVal StartPosition As Long = -1, _
              Optional ByVal FinishPosition As Long = -1) As Variant
    If IsEmpty(ValueArray) Then Exit Function
    Dim StartIndex As Long
    Dim FinishIndex As Long
    StartIndex = IIf(StartPosition = -1, LBound(ValueArray, 1), StartPosition)
    FinishIndex = IIf(FinishPosition = -1, UBound(ValueArray, 1), FinishPosition)
    
    '''' Prepare an array of indecies to be sorted.
    '''' Each element in ValueIndex says that in the position "i" of the sorted array
    '''' should be an element with index ValueIndex(i) from the original unordered array.
    Dim ValueIndex() As Long
    ReDim ValueIndex(StartIndex To FinishIndex)
    Dim Index As Long
    For Index = StartIndex To FinishIndex
        ValueIndex(Index) = Index
    Next Index
    
    Randomize

    '''' Sort array of indecies
    VectorSort ValueArray, ValueIndex, StartIndex, FinishIndex
    
    '''' Use sorted ValueIndex array to order original array in place
    Dim Buffer As Variant
    Dim ValuePointer As Long
    Dim ValuePointerNext As Long
    Dim ValuePointerBuffer As Long
    ValuePointerNext = FinishIndex
    
    '''' The original unordered array may be represented by one or several rings, such that
    '''' within each ring elements change positions between themselves. Starting from a given
    '''' end/direction locate the first element that is out of order and order the elements
    '''' in the ValueArray and ValueIndex, tracing the corresponding ring. If that ring does
    '''' not include all elements, continue scanning the ValueIndex array and find the first
    '''' remaining element that is out of order and repeat cycle, until both arrays are oredered.
    Do While ValuePointerNext > StartIndex
        '''' Start from the last element and go backwards. Find the first element that is out of order
        ValuePointer = ValuePointerNext
        Do While (ValueIndex(ValuePointer) = ValuePointer) And (ValuePointer > StartIndex)
            ValuePointer = ValuePointer - 1
        Loop
        If ValuePointer = StartIndex Then
            Exit Do
        Else
            ValuePointerNext = ValuePointer - 1
        End If
        
        '''' Order the current ring
        Buffer = ValueArray(ValuePointer)
        Do While ValueIndex(ValuePointer) < ValuePointerNext + 1
            ValueArray(ValuePointer) = ValueArray(ValueIndex(ValuePointer))
            ValuePointerBuffer = ValueIndex(ValuePointer)
            ValueIndex(ValuePointer) = ValuePointer
            ValuePointer = ValuePointerBuffer
        Loop
        ValueArray(ValuePointer) = Buffer
        ValueIndex(ValuePointer) = ValuePointer
    Loop
    Vector = ValueArray
End Function


'@Description "Performs recursive Q-sorting"
Private Sub VectorSort(ByRef ValueArray As Variant, _
                               ByRef ValueIndex As Variant, _
                      Optional ByVal StartIndex As Long = -1, _
                      Optional ByVal FinishIndex As Long = -1)
Attribute VectorSort.VB_Description = "Performs recursive Q-sorting"
    Dim PivotValue As Variant
    Dim PivotIndex As Long
    Dim LeftIndex As Long
    Dim RightIndex As Long
    Dim Buffer As Long

    If StartIndex >= FinishIndex Then Exit Sub
    
    PivotIndex = StartIndex + CLng(Round(Rnd * (FinishIndex - StartIndex)))
    PivotValue = ValueArray(ValueIndex(PivotIndex))
    LeftIndex = StartIndex
    RightIndex = FinishIndex
                        
    Do While LeftIndex < RightIndex
        Do While (ValueArray(ValueIndex(RightIndex)) > PivotValue) And (LeftIndex < RightIndex)
            RightIndex = RightIndex - 1
        Loop
        Do While (ValueArray(ValueIndex(LeftIndex)) < PivotValue) And (LeftIndex < RightIndex)
             LeftIndex = LeftIndex + 1
        Loop
        If LeftIndex < RightIndex Then
            Buffer = ValueIndex(RightIndex)
            ValueIndex(RightIndex) = ValueIndex(LeftIndex)
            ValueIndex(LeftIndex) = Buffer
            RightIndex = RightIndex - 1
            LeftIndex = LeftIndex + 1
        End If
    Loop
    
    '''' Handle edge cases
    If LeftIndex > RightIndex Then
        RightIndex = RightIndex + 1
        LeftIndex = LeftIndex - 1
    ElseIf LeftIndex = RightIndex Then
        If RightIndex = FinishIndex Then
            LeftIndex = LeftIndex - 1
        ElseIf LeftIndex = StartIndex Then
            RightIndex = RightIndex + 1
        ElseIf ValueArray(ValueIndex(RightIndex)) >= PivotValue Then
            LeftIndex = LeftIndex - 1
        Else
            RightIndex = RightIndex + 1
        End If
    End If
          
    If StartIndex < LeftIndex Then
        VectorSort ValueArray, ValueIndex, StartIndex, LeftIndex
    End If
    If RightIndex < FinishIndex Then
        VectorSort ValueArray, ValueIndex, RightIndex, FinishIndex
    End If
End Sub
