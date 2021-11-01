Attribute VB_Name = "ProcessToolsDemo"
'@Folder "Process"
'@IgnoreModule IntegerDataType
Option Explicit

Private Const PROCESS_VM_READ As Long = &H10&
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&

Private Const SQLITE_SIGNATURE_DW0_LE As Long = &H694C5153
Private Const SQLITE_SIGNATURE_DW1_LE As Long = &H66206574
Private Const SQLITE_SIGNATURE_DW2_LE As Long = &H616D726F
Private Const SQLITE_SIGNATURE_DW3_LE As Long = &H332074
    
Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    #If VBA7 Then
        lpMinimumApplicationAddress As LongPtr
        lpMaximumApplicationAddress As LongPtr
        dwActiveProcessorMask As LongPtr
    #Else
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
    #End If
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type MEMORY_BASIC_INFORMATION
    #If VBA7 Then
        BaseAddress As LongPtr
        AllocationBase As LongPtr
        AllocationProtect As Long
        RegionSize As LongPtr
    #Else
        BaseAddress As Long
        AllocationBase As Long
        AllocationProtect As Long
        RegionSize As Long
    #End If
    State As Long
    Protect As Long
    lType As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
    Private Declare PtrSafe Function VirtualQueryEx Lib "kernel32" Alias "VirtualQueryEx" ( _
        ByVal hProcess As LongPtr, ByVal lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" ( _
        ByVal hProcess As LongPtr, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As LongPtr, ByRef lpNumberOfBytesWritten As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
    Private Declare Function VirtualQueryEx Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
    Private Declare Function ReadProcessMemory Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
#End If


Public Sub GetSQLiteMem()
    Dim PID As Long
    PID = GetCurrentProcessId()
    
    Dim SystemInfo As SYSTEM_INFO
    GetSystemInfo SystemInfo
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, False, PID)
    
    Dim dbcMem As SQLiteCConnection
    Set dbcMem = FixObjC.GetDBCMemITRB
''''    Dim dbcTmp As SQLiteCConnection
''''    Set dbcTmp = FixObjC.GetDBCTempFuncWithData
    
''''    Dim PackedHeader As SQLiteCHeaderPacked
''''    Dim PackedHeaderCopy(0 To 1023) As Long
''''    '@Ignore IntegerDataType
''''    Dim FileHandle As Integer
''''    FileHandle = FreeFile()
''''    Open dbcTmp.DbPathName For Random As FileHandle Len = Len(PackedHeader)
''''    Get FileHandle, 1, PackedHeader
''''    Close FileHandle
    
''''    RtlMoveMemory VarPtr(PackedHeaderCopy(0)), VarPtr(PackedHeader), Len(PackedHeader)
    
    #If VBA7 Then
        Dim DbAddress As LongPtr
        Dim BaseAddress As LongPtr
    #Else
        Dim DbAddress As Long
        Dim BaseAddress  As Long
    #End If
    
    Dim MemoryBlockSize As Long
    Dim BufferSize As Long
    Dim Buffer() As Long
    
    Dim ElementIndex As Long
    Dim ElementTop As Long
    Dim MemoryInfo As MEMORY_BASIC_INFORMATION
    
    Dim PackedHeader As SQLiteCHeaderPacked
    Dim dbh As SQLiteCHeader
    Set dbh = SQLiteCHeader.Create()
    Dim Result As Long
    Dim ByteCount As Long
    BufferSize = 0
    BaseAddress = SystemInfo.lpMinimumApplicationAddress
    Do While BaseAddress < SystemInfo.lpMaximumApplicationAddress
        ByteCount = VirtualQueryEx(hProcess, BaseAddress, MemoryInfo, Len(MemoryInfo))
        Debug.Assert ByteCount = Len(MemoryInfo)
        MemoryBlockSize = MemoryInfo.RegionSize
        If MemoryBlockSize > BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        End If
        
        Result = ReadProcessMemory(hProcess, BaseAddress, Buffer(0), MemoryBlockSize, ByteCount)
        If Result <> 0 Then
            ElementTop = ByteCount / 4 - 1
            For ElementIndex = 0 To ElementTop
                If Buffer(ElementIndex) = SQLITE_SIGNATURE_DW0_LE Then
                    If Buffer(ElementIndex + 1) = SQLITE_SIGNATURE_DW1_LE And _
                       Buffer(ElementIndex + 2) = SQLITE_SIGNATURE_DW2_LE And _
                       Buffer(ElementIndex + 3) = SQLITE_SIGNATURE_DW3_LE And _
                       Buffer(ElementIndex + 17) = 0 And Buffer(ElementIndex + 18) = 0 And _
                       Buffer(ElementIndex + 19) = 0 And Buffer(ElementIndex + 20) = 0 And _
                       Buffer(ElementIndex + 21) = 0 _
                       Then
                        DbAddress = BaseAddress + ElementIndex * Len(Buffer(0))
                        RtlMoveMemory VarPtr(PackedHeader), DbAddress, Len(PackedHeader)
                        dbh.UnpackHeader PackedHeader
                        If (Buffer(ElementIndex + 5) And &HF0FFFFFF) = 0 Then GoTo SIGNATURE_MATCH:
                    End If
                End If
            Next ElementIndex
        End If
        BaseAddress = BaseAddress + MemoryBlockSize
    Loop

SIGNATURE_MATCH:
    Result = CloseHandle(hProcess)
End Sub
