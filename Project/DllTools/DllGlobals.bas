Attribute VB_Name = "DllGlobals"
'@Folder "DllTools"
Option Explicit

Public Const ERROR_BAD_EXE_FORMAT As Long = 193
Public Const LoadingDllErr As Long = 48
Public Const DT_ICU_V As String = "68"

'''' https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comtypes.callconv
Public Enum CALLCONV
    CC_CDECL = 1&
    CC_MSCPASCAL = 2&
    CC_PASCAL = 2&
    CC_MACPASCAL = 3&
    CC_STDCALL = 4&
    CC_RESERVED = 5&
    CC_SYSCALL = 6&
    CC_MPWCDECL = 7&
    CC_MPWPASCAL = 8&
    CC_MAX = 9&
End Enum

Public Enum HRESULT
    S_OK = &H0&                 '''' Operation successful
    E_ABORT = &H80004004        '''' Operation aborted
    E_ACCESSDENIED = &H80070005 '''' General access denied error
    E_FAIL = &H80004005         '''' Unspecified failure
    E_HANDLE = &H80070006       '''' Handle that is not valid
    E_INVALIDARG = &H80070057   '''' One or more arguments are not valid
    E_NOINTERFACE = &H80004002  '''' No such interface supported
    E_NOTIMPL = &H80004001      '''' Not implemented
    E_OUTOFMEMORY = &H8007000E  '''' Failed to allocate necessary memory
    E_POINTER = &H80004003      '''' Pointer that is not valid
    E_UNEXPECTED = &H8000FFFF   '''' Unexpected failure
End Enum

Public Enum DllLoadStatus
    LOAD_OK = -1
    LOAD_FAIL = 0
    LOAD_ALREADY_LOADED = 1
End Enum
