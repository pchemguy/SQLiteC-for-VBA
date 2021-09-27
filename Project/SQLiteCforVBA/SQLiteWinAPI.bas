Attribute VB_Name = "SQLiteWinAPI"
'@Folder "SQLiteCforVBA"
Option Explicit

Public Const CP_UTF8 As Long = 65001
Public Const ErrorLoadingDLL As Long = 48
Private Const JULIANDAY_OFFSET As Double = 2415018.5

#If VBA7 Then

Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As LongPtr, ByVal pSource As LongPtr, ByVal length As Long)
Public Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr, ByVal cchCount As Long) As LongPtr
Public Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr) As LongPtr
Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As LongPtr) As Long
Public Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As LongPtr) As LongPtr
Public Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As LongPtr) As Long
Public Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Public Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Public Declare PtrSafe Function SetDllDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

#Else

Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal length As Long)
Public Declare Function lstrcpynW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long, ByVal cchCount As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal pwsString As Long) As Long
Public Declare Function SysAllocString Lib "OleAut32" (ByRef pwsString As Long) As Long
Public Declare Function SysStringLen Lib "OleAut32" (ByVal bstrString As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function SetDllDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

#End If


#If Win64 Then
    Private hSQLiteLibrary As LongPtr
#Else
    Private hSQLiteLibrary As Long
#End If


Public Function LoadLib(Optional ByVal LibDir As String = vbNullString) As Long
    ' SetDllDirectory can also be used
    Dim fso As New Scripting.FileSystemObject
    Dim PathName As String
    PathName = fso.GetAbsolutePathName(IIf(Len(LibDir) > 0, LibDir, _
        ThisWorkbook.Path)) & Application.PathSeparator & "SQLite3.dll"
        
    If hSQLiteLibrary <> 0 Then
        Err.Raise ErrorLoadingDLL, "SQLiteWinAPI", "The library already loaded"
    End If
    
    hSQLiteLibrary = LoadLibrary(PathName)
    If hSQLiteLibrary = 0 Then
        Debug.Print "LoadLib Error Loading " & PathName & ":", Err.LastDllError
        LoadLib = SQLITE_ERROR
    Else
        LoadLib = SQLITE_OK
    End If
End Function


Public Sub FreeLib()
    If hSQLiteLibrary = 0 Then
        Err.Raise ErrorLoadingDLL, "SQLiteWinAPI", "The library is not loaded"
    End If
    
    Dim Result As Long
    Result = FreeLibrary(hSQLiteLibrary)
    If Result <> 0 Then
        hSQLiteLibrary = 0
    Else
        Debug.Print "SQLite3Free Error Freeing SQLite3.dll:", Result, Err.LastDllError
    End If
End Sub


' String Helpers
#If VB7 Then
Public Function UTF8PtrToString(ByVal UTF8StrPtr As LongPtr) As String
#Else
Public Function UTF8PtrToString(ByVal UTF8StrPtr As Long) As String
#End If
    Dim Buffer As String
    Dim Size As Long
    Dim RetVal As Long
    Dim Result As String
    
    Size = MultiByteToWideChar(CP_UTF8, 0, UTF8StrPtr, -1, 0, 0)
    ' Size includes the terminating null character
    If Size <= 1 Then
        UTF8PtrToString = vbNullString
        Exit Function
    End If
    
    Result = String(Size - 1, " ") ' and a termintating null char.
    RetVal = MultiByteToWideChar(CP_UTF8, 0, UTF8StrPtr, -1, StrPtr(Result), Size)
    If RetVal = 0 Then
        Debug.Print "Utf8PtrToString Error:", Err.LastDllError
        Debug.Assert RetVal > 0
        Exit Function
    End If
    UTF8PtrToString = Result
End Function


Public Function StringToUtf8Bytes(ByVal Text As String) As Variant
    Dim Size As Long
    Dim RetVal As Long
    Dim Buffer() As Byte
    
    Size = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
    If Size = 0 Then
        Exit Function
    End If
    
    ReDim Buffer(Size)
    RetVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, VarPtr(Buffer(0)), Size, 0, 0)
    If RetVal = 0 Then
        Debug.Print "StringToUtf8Bytes Error:", Err.LastDllError
        Exit Function
    End If
    StringToUtf8Bytes = Buffer
End Function


#If VBA7 Then
Public Function UTF16PtrToString(ByVal UTF16StrPtr As LongPtr) As String
#Else
Public Function UTF16PtrToString(ByVal UTF16StrPtr As Long) As String
#End If
    Dim StrLen As Long
    StrLen = lstrlenW(UTF16StrPtr)
    UTF16PtrToString = String(StrLen, " ")
    lstrcpynW StrPtr(UTF16PtrToString), UTF16StrPtr, StrLen
End Function


' Date Helpers
Public Function ToJulianDay(oleDate As Date) As Double
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function


Public Function FromJulianDay(julianDay As Double) As Date
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function

