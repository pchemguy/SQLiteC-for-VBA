Attribute VB_Name = "ProcessToolsDemo"
'@Folder "SQLiteDBdev.Process"
'@IgnoreModule IntegerDataType, UseMeaningfulName, HungarianNotation
'@IgnoreModule
'
'''' This module contains code aimed at locating a certain data structure in
'''' memory (in this case, an in-memory SQLite db is targeted). Here, the goal
'''' is to locate an in-memory SQLite db header, parse it, determine db size,
'''' and dump it to a file. Routines are drafted for searching both own memory
'''' address space and another process's memory. When searching another process's
'''' address space, care must be taken to keep references pointing to foreign and
'''' own address space separate. One of the reason for creating this draft was an
'''' attempt to pull the SQLite database (db/api.db) from memory of running
'''' Windows API Viewer for MS Excel. The application is packed with BoxedApp app,
'''' which implements a virtual file system. An attempt to pull original db failed.
'''' While db file candidate was located and dumped successfully and the location
'''' was overall correct based on the contents, the resulting db file was not usable.
'''' Apparently, BoxedApp employs some anti-reverese engineering protection. Further
'''' investigation is necessary to sort it out. The db file was also dumped
'''' independently via HxD memory access feature with the same result. Because this
'''' use case for functionality developed in this module could not be completed,
'''' this module is left as is (working draft) and no time invested in creating a
'''' memory manager class.

Option Explicit

Private Const PROCESS_VM_READ As Long = &H10&
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const PROCESS_VM_OPERATION As Long = &H8&

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

Private Enum MEMORY_BASIC_INFO_TYPE
    MEMORY_BASIC_INFO_TYPE_MEM_PRIVATE = &H20000
    MEMORY_BASIC_INFO_TYPE_MEM_MAPPED = &H40000
    MEMORY_BASIC_INFO_TYPE_MEM_IMAGE = &H1000000
End Enum

Private Enum MEMORY_BASIC_INFO_STATE
    MEMORY_BASIC_INFO_STATE_MEM_COMMIT = &H1000&
    MEMORY_BASIC_INFO_STATE_MEM_RESERVE = &H2000&
    MEMORY_BASIC_INFO_STATE_MEM_FREE = &H10000
End Enum

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
    State As MEMORY_BASIC_INFO_STATE
    Protect As Long
    lType As MEMORY_BASIC_INFO_TYPE
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    #If VBA7 Then
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    #Else
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    #End If
End Type

Private Type PROCESS_INFORMATION
    #If VBA7 Then
        hProcess As LongPtr
        hThread As LongPtr
    #Else
        hProcess As Long
        hThread As Long
    #End If
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    #If VBA7 Then
        lpSecurityDescriptor As LongPtr
    #Else
        lpSecurityDescriptor As Long
    #End If
    bInheritHandle As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function EnumProcesses Lib "psapi" (ByRef lpidProcess As LongPtr, ByVal cb As Long, ByRef cbNeeded As LongPtr) As Long
    Private Declare PtrSafe Function GetMappedFileNameA Lib "psapi" (ByVal hProcess As LongPtr, ByVal lpv As LongPtr, _
                                                                     ByVal lpFilename As String, ByVal nSize As Long) As Long
    
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByRef lpEnvironment As Any, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As LongPtr
    
    Private Declare PtrSafe Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
    Private Declare PtrSafe Function VirtualQueryEx Lib "kernel32" ( _
        ByVal hProcess As LongPtr, ByVal lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReadProcessMemory Lib "kernel32" ( _
        ByVal hProcess As LongPtr, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As LongPtr, ByRef lpNumberOfBytesWritten As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal Length As Long)
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function EnumProcesses Lib "psapi" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Private Declare Function GetMappedFileNameA Lib "psapi" (ByVal hProcess As Long, ByVal lpv As Long, _
                                                             ByVal lpFilename As String, ByVal nSize As Long) As Long
    
    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpCommandLine As String, _
        ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, _
        ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByRef lpEnvironment As Any, _
        ByVal lpCurrentDirectory As String, _
        ByRef lpStartupInfo As STARTUPINFO, _
        ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

    Private Declare Sub GetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)
    Private Declare Function VirtualQueryEx Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
    Private Declare Function ReadProcessMemory Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
#End If


Public Sub GetSQLiteMemTest()
    Dim UserVersion As Long
    UserVersion = &HAABBCCDD
    Dim AppId As Long
    AppId = &H11223344
    Dim DummyBuffer() As Long
    DummyBuffer = SQLiteCHeader.GenBlankDb( _
        UserVersion:=SQLiteCHeader.SLong32bFlipBytes(UserVersion), _
              ApplicationId:=SQLiteCHeader.SLong32bFlipBytes(AppId))
    Debug.Print VarPtr(DummyBuffer(0))
    
    Dim PID As Long
    PID = GetCurrentProcessId()
    
    Const MAX_BLOCK_SIZE As Long = 1024& * 1024&
    Dim SystemInfo As SYSTEM_INFO
    GetSystemInfo SystemInfo
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, False, PID)
    
    Dim dbcMem As SQLiteCConnection
    Set dbcMem = FixObjC.GetDBCMemFuncWithData
    
    #If VBA7 Then
        Dim DbAddress As LongPtr
        Dim BaseAddress As LongPtr
        Dim BufferSize As LongPtr
        Dim MemoryBlockSize As LongPtr
        Dim ByteCount As LongPtr
    #Else
        Dim DbAddress As Long
        Dim BaseAddress  As Long
        Dim BufferSize As Long
        Dim MemoryBlockSize As Long
        Dim ByteCount As Long
    #End If
    
    Dim Buffer() As Long
    
    Dim ElementIndex As Long
    Dim ElementTop As Long
    Dim MemoryInfo As MEMORY_BASIC_INFORMATION
    
    Dim Result As Long
    BufferSize = 0
    
    BaseAddress = SystemInfo.lpMinimumApplicationAddress
    Do While BaseAddress < SystemInfo.lpMaximumApplicationAddress
        ByteCount = VirtualQueryEx(hProcess, BaseAddress, MemoryInfo, Len(MemoryInfo))
        Debug.Assert ByteCount = Len(MemoryInfo)
        MemoryBlockSize = MemoryInfo.RegionSize
        If MemoryBlockSize > MAX_BLOCK_SIZE Then
            MemoryBlockSize = MAX_BLOCK_SIZE
        End If
        If MemoryBlockSize > BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        ElseIf MemoryBlockSize * 16 < BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        End If
        BaseAddress = MemoryInfo.BaseAddress
        Result = ReadProcessMemory(hProcess, BaseAddress, Buffer(0), MemoryBlockSize, ByteCount)
        
        If Result <> 0 Then
            ElementTop = ByteCount / 4 - 1
            For ElementIndex = 0 To ElementTop
                If Buffer(ElementIndex) = SQLITE_SIGNATURE_DW0_LE Then
                    If Buffer(ElementIndex + 1) = SQLITE_SIGNATURE_DW1_LE And _
                       Buffer(ElementIndex + 2) = SQLITE_SIGNATURE_DW2_LE And _
                       Buffer(ElementIndex + 3) = SQLITE_SIGNATURE_DW3_LE And _
                       Buffer(ElementIndex + 15) = UserVersion And _
                       Buffer(ElementIndex + 17) = AppId And _
                       Buffer(ElementIndex + 18) = 0 And Buffer(ElementIndex + 19) = 0 And _
                       Buffer(ElementIndex + 20) = 0 And Buffer(ElementIndex + 21) = 0 And _
                       Buffer(ElementIndex + 22) = 0 _
                       Then
                        DbAddress = BaseAddress + ElementIndex * Len(Buffer(0))
                        DumpMemDb DbAddress
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


Private Sub DumpMemDb(ByVal DbAddress As Variant)
    Dim dbh As SQLiteCHeader
    Set dbh = SQLiteCHeader.Create()
    Dim PackedHeader As SQLiteCHeaderPacked
    RtlMoveMemory VarPtr(PackedHeader), DbAddress, Len(PackedHeader)
    dbh.UnpackHeader PackedHeader
    
    If dbh.Header.DbFilePageCount <= 0 Or dbh.Header.DbFilePageCount > 100000 Then
        Debug.Print "Bad image at " & CStr(DbAddress)
        Exit Sub
    End If
    
    Dim DbSize As Long
    DbSize = dbh.Header.PageSizeInBytes * dbh.Header.DbFilePageCount
    Dim DbData() As Byte
    ReDim DbData(0 To DbSize - 1)
    RtlMoveMemory VarPtr(DbData(0)), DbAddress, DbSize
    
    Dim DbPathName As String
    DbPathName = FixObjC.RandomTempFileName("-----" & CStr(DbAddress) & ".db")
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open DbPathName For Binary Access Read Write As FileHandle
    Put FileHandle, , DbData
    Close FileHandle
End Sub


Private Sub SaveBlankDb()
    Dim Buffer() As Long
    Buffer = SQLiteCHeader.GenBlankDb(UserVersion:=&HDDCCBBAA, ApplicationId:=&H77665544)
    Dim DbPathName As String
    DbPathName = FixObjC.RandomTempFileName("-----.db")
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open DbPathName For Binary Access Read Write As FileHandle
    Put FileHandle, , Buffer
    Close FileHandle
End Sub


Private Sub SaveArray()
    Dim Buffer(0 To 2) As Long
    Buffer(0) = &H12345678
    Buffer(1) = &H87654321
    Buffer(2) = &HABCDEF00
    Dim DbPathName As String
    DbPathName = FixObjC.RandomTempFileName("-----.db")
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open DbPathName For Binary Access Read Write As FileHandle
    Put FileHandle, , Buffer
    Close FileHandle
End Sub


Private Sub RunGetSQLiteMem()
    Dim ChildPath As String
    ChildPath = Environ("USERPROFILE") & "\Downloads\WindowsAPIViewer\x86\WinAPIExcelp.exe"
    Dim PID As Long
    PID = Shell(ChildPath, vbMinimizedNoFocus)
    GetSQLiteMem PID
'    #If VBA7 Then
'        Dim hProcess As LongPtr
'    #Else
'        Dim hProcess As Long
'    #End If
'    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, False, PID)
'
'    Dim Result As Long
'    Result = CloseHandle(hProcess)
End Sub


Public Sub GetSQLiteMem(Optional ByVal ProcessID As Long = 0)
    Dim PID As Long
    PID = IIf(ProcessID = 0, GetCurrentProcessId(), ProcessID)
    
    Const MAX_BLOCK_SIZE As Long = 1024& * 1024&
    Dim SystemInfo As SYSTEM_INFO
    GetSystemInfo SystemInfo
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION, False, PID)
        
    #If VBA7 Then
        Dim BaseAddress As LongPtr
        Dim MemoryBlockSize As LongPtr
        Dim MemoryBlockSizeOld As LongPtr
        Dim BufferSize As LongPtr
        Dim ByteCount As LongPtr
    #Else
        Dim BaseAddress  As Long
        Dim MemoryBlockSize As Long
        Dim MemoryBlockSizeOld As Long
        Dim BufferSize As Long
        Dim ByteCount As Long
    #End If
    
    Dim Buffer() As Long
    
    Dim ElementIndex As Long
    Dim ElementTop As Long
    Dim MemoryInfo As MEMORY_BASIC_INFORMATION
    
    Dim Result As Long
    BufferSize = 0
    
    BaseAddress = SystemInfo.lpMinimumApplicationAddress
    Do While BaseAddress < SystemInfo.lpMaximumApplicationAddress
        ByteCount = VirtualQueryEx(hProcess, BaseAddress, MemoryInfo, Len(MemoryInfo))
        Debug.Assert ByteCount = Len(MemoryInfo)
        MemoryBlockSize = MemoryInfo.RegionSize
        If MemoryBlockSize > MAX_BLOCK_SIZE Then
            MemoryBlockSize = MAX_BLOCK_SIZE
        End If
        If MemoryBlockSize > BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        ElseIf MemoryBlockSize * 16 < BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        End If
        BaseAddress = MemoryInfo.BaseAddress
        Result = ReadProcessMemory(hProcess, BaseAddress, Buffer(0), MemoryBlockSize, ByteCount)
        
        If Result <> 0 Then
            ElementTop = ByteCount / 4 - 1
            For ElementIndex = 0 To ElementTop
                If Buffer(ElementIndex) = SQLITE_SIGNATURE_DW0_LE Then
                    If Buffer(ElementIndex + 1) = SQLITE_SIGNATURE_DW1_LE And _
                       Buffer(ElementIndex + 2) = SQLITE_SIGNATURE_DW2_LE And _
                       Buffer(ElementIndex + 3) = SQLITE_SIGNATURE_DW3_LE Then
                        GoSub COLLECT_DB:
                        'DbAddress = VarPtr(Buffer(0)) + ElementIndex * Len(Buffer(0))
                        'Debug.Print CStr(DbAddress)
                        'DumpMemDb DbAddress
                    End If
                End If
            Next ElementIndex
        End If
        BaseAddress = BaseAddress + MemoryBlockSize
    Loop
    Result = CloseHandle(hProcess)
    
    Exit Sub

COLLECT_DB:
    Dim PackedHeader As SQLiteCHeaderPacked
    Dim HeaderSize As Long
    HeaderSize = Len(PackedHeader)
    Dim HeaderBuffer() As Byte
    ReDim HeaderBuffer(0 To HeaderSize - 1)
    
    #If VBA7 Then
        Dim DbAddressRemote As LongPtr
        Dim DbAddressLocal As LongPtr
        Dim DbDataCursor As LongPtr
        Dim BufferCursor As LongPtr
        Dim BaseAddressOld As LongPtr
    #Else
        Dim DbAddressRemote As Long
        Dim DbAddressLocal As Long
        Dim DbDataCursor As Long
        Dim BufferCursor As Long
        Dim BaseAddressOld As Long
    #End If
    
    BaseAddressOld = BaseAddress
    Dim ElementIndexOld As Long
    ElementIndexOld = ElementIndex
    MemoryBlockSizeOld = MemoryBlockSize
    
    DbAddressRemote = BaseAddress + ElementIndex * Len(Buffer(0))
    DbAddressLocal = VarPtr(Buffer(0)) + ElementIndex * Len(Buffer(0))
    RtlMoveMemory VarPtr(HeaderBuffer(0)), DbAddressLocal, HeaderSize
    Dim dbh As SQLiteCHeader
    Set dbh = SQLiteCHeader.Create(vbNullString)
    PackedHeader = dbh.PackedHeaderFromBytes(HeaderBuffer)
    dbh.UnpackHeader PackedHeader
    Dim DbSize As Long
    DbSize = dbh.Header.DbFilePageCount * dbh.Header.PageSizeInBytes
    If DbSize = 0 Then
        Debug.Print CStr(DbAddressRemote) & " - bad db image."
        Return
    End If
    Debug.Print CStr(DbAddressRemote), CStr(DbSize)
    
    Dim DbData() As Byte
    ReDim DbData(0 To DbSize - 1)
    DbDataCursor = VarPtr(DbData(0))
    BufferCursor = VarPtr(Buffer(0)) + ElementIndex * Len(Buffer(0))
    
    Dim DbTailSize As Long
    DbTailSize = DbSize
    Dim BufferTailSize As Long
    BufferTailSize = CLng(ByteCount) - ElementIndex * Len(Buffer(0))
    Do While BufferTailSize < DbTailSize
        RtlMoveMemory DbDataCursor, BufferCursor, BufferTailSize
        DbDataCursor = DbDataCursor + BufferTailSize
        DbTailSize = DbTailSize - BufferTailSize
        BaseAddress = BaseAddress + MemoryBlockSize

        ByteCount = VirtualQueryEx(hProcess, BaseAddress, MemoryInfo, Len(MemoryInfo))
        Debug.Assert ByteCount = Len(MemoryInfo)
        MemoryBlockSize = MemoryInfo.RegionSize
        If MemoryBlockSize > MAX_BLOCK_SIZE Then
            MemoryBlockSize = MAX_BLOCK_SIZE
        End If
        If MemoryBlockSize > BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        ElseIf MemoryBlockSize * 16 < BufferSize Then
            BufferSize = MemoryBlockSize
            ReDim Buffer(0 To BufferSize / Len(Buffer(0)) - 1)
        End If
        BaseAddress = MemoryInfo.BaseAddress
        Result = ReadProcessMemory(hProcess, BaseAddress, Buffer(0), MemoryBlockSize, ByteCount)
        If ByteCount < MemoryBlockSize And ByteCount < DbTailSize Then
            BaseAddress = BaseAddressOld
            MemoryBlockSize = MemoryBlockSizeOld
            ElementIndex = ElementIndexOld + Len(PackedHeader) / Len(Buffer(0))
            Return
        End If
        ElementIndex = 0
        BufferTailSize = CLng(ByteCount)
        BufferCursor = VarPtr(Buffer(0))
    Loop
    If DbTailSize > 0 Then
        RtlMoveMemory DbDataCursor, BufferCursor, DbTailSize
        ElementIndex = ElementIndex + DbTailSize \ Len(Buffer(0))
    End If
    BaseAddress = BaseAddressOld
    MemoryBlockSize = MemoryBlockSizeOld
    ElementIndex = ElementIndexOld + Len(PackedHeader) / Len(Buffer(0))
    
    GoSub DUMP_DB:
    Return
    
    
DUMP_DB:
    Dim DbPathName As String
    DbPathName = FixObjC.RandomTempFileName("-----" & CStr(DbAddressRemote) & ".db")
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open DbPathName For Binary Access Read Write As FileHandle
    Put FileHandle, , DbData
    Close FileHandle
    Return
End Sub


''''    Dim dbcTmp As SQLiteCConnection
''''    Set dbcTmp = FixObjC.GetDBCTempFuncWithData
    
''''    Dim PackedHeader As SQLiteCHeaderPacked
''''    Dim PackedHeaderCopy(0 To 1023) As Long
''''    '@Ignore IntegerDataType
''''    Dim FileHandle As Integer
''''    FileHandle = FreeFile()
''''    Open dbcTmp.DbPathName For Random As FileHandle Len = Len(PackedHeader)
''''    Get FileHandle, 1,   PackedHeader
''''    Close FileHandle
    
''''    RtlMoveMemory VarPtr(PackedHeaderCopy(0)), VarPtr(PackedHeader), Len(PackedHeader)


Private Sub RunChildCreateProcess()
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    
    Dim CommandLine As String
    CommandLine = Environ("USERPROFILE") & "\Downloads\WindowsAPIViewer\x86\WinAPIExcelp.exe"

    Dim SecAttrProc As SECURITY_ATTRIBUTES
    SecAttrProc.nLength = Len(SecAttrProc)
    Dim SecAttrThr As SECURITY_ATTRIBUTES
    SecAttrThr.nLength = Len(SecAttrThr)

    Dim StartInfo As STARTUPINFO ' : startInfo.cb = len(startInfo)
    Dim ProcInfo As PROCESS_INFORMATION

    hProcess = CreateProcess( _
        lpApplicationName:=CommandLine, _
        lpCommandLine:=vbNullString, _
        lpProcessAttributes:=SecAttrProc, _
        lpThreadAttributes:=SecAttrThr, _
        bInheritHandles:=False, _
        dwCreationFlags:=0, _
        lpEnvironment:=0, _
        lpCurrentDirectory:=vbNullString, _
        lpStartupInfo:=StartInfo, _
        lpProcessInformation:=ProcInfo)
        
    CloseHandle ProcInfo.hProcess
End Sub


Private Sub RunSurveyMem()
    Dim ChildPath As String
    ChildPath = Environ("USERPROFILE") & "\Downloads\WindowsAPIViewer\x86\WinAPIExcelp.exe"
    Dim PID As Long
    PID = Shell(ChildPath, vbMinimizedNoFocus)
    SurveyMem PID
End Sub


Public Sub SurveyMem(Optional ByVal ProcessID As Long = 0)
    Dim MemInfoTypeMap As Scripting.Dictionary
    Set MemInfoTypeMap = New Scripting.Dictionary
    With MemInfoTypeMap
        .CompareMode = TextCompare
        .Item(MEMORY_BASIC_INFO_TYPE_MEM_PRIVATE) = "PRIVATE"
        .Item(MEMORY_BASIC_INFO_TYPE_MEM_MAPPED) = "MAPPED"
        .Item(MEMORY_BASIC_INFO_TYPE_MEM_IMAGE) = "IMAGE"
    End With
    Dim MemInfoStateMap As Scripting.Dictionary
    Set MemInfoStateMap = New Scripting.Dictionary
    With MemInfoStateMap
        .CompareMode = TextCompare
        .Item(MEMORY_BASIC_INFO_STATE_MEM_COMMIT) = "COMMIT"
        .Item(MEMORY_BASIC_INFO_STATE_MEM_RESERVE) = "RESERVE"
        .Item(MEMORY_BASIC_INFO_STATE_MEM_FREE) = "FREE"
    End With
    
    Dim Result As Long
    Dim PID As Long
    PID = IIf(ProcessID = 0, GetCurrentProcessId(), ProcessID)
    
    Dim SystemInfo As SYSTEM_INFO
    GetSystemInfo SystemInfo
    #If VBA7 Then
        Dim hProcess As LongPtr
    #Else
        Dim hProcess As Long
    #End If
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION, False, PID)
        
    #If VBA7 Then
        Dim BaseAddress As LongPtr
        Dim ByteCount As LongPtr
        Dim MemoryBlockSize As LongPtr
    #Else
        Dim BaseAddress  As Long
        Dim ByteCount As Long
        Dim MemoryBlockSize As Long
    #End If
    
    Const MAX_PATH_LENGTH As Long = 512
    Dim MemInfoSummary As String
    Dim FilePathName As String * MAX_PATH_LENGTH
    Dim Blank As String
    Blank = String(Len(FilePathName), vbNullChar)
    Dim MemoryInfo As MEMORY_BASIC_INFORMATION
    BaseAddress = SystemInfo.lpMinimumApplicationAddress
    Do While BaseAddress < SystemInfo.lpMaximumApplicationAddress
        ByteCount = VirtualQueryEx(hProcess, BaseAddress, MemoryInfo, Len(MemoryInfo))
        Debug.Assert ByteCount = Len(MemoryInfo)
        MemoryBlockSize = MemoryInfo.RegionSize
        BaseAddress = MemoryInfo.BaseAddress
        MemInfoSummary = "Address: " & CStr(BaseAddress) & _
                           " Type: " & CStr(MemInfoTypeMap(MemoryInfo.lType)) & _
                         ". State: " & CStr(MemInfoStateMap(MemoryInfo.State))
        FilePathName = Blank
        Result = GetMappedFileNameA(hProcess, BaseAddress, FilePathName, MAX_PATH_LENGTH)
        If Result > 0 Then
            MemInfoSummary = MemInfoSummary & ". " & vbNewLine & vbTab & _
                                              "Mapped file: " & CStr(FilePathName)
            Debug.Print MemInfoSummary
        End If
                                      
        BaseAddress = BaseAddress + MemoryBlockSize
    Loop
    Result = CloseHandle(hProcess)
End Sub


