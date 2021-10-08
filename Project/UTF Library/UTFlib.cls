VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTFlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Performs UTF text conversion"
'@Folder "UTF Library"
'@ModuleDescription "Performs UTF text conversion"
'@PredeclaredId
Option Explicit

Private Const CP_UTF8 As Long = 65001

#If WIN64 Then
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pwsString As LongPtr) As Long
Private Declare PtrSafe Function lstrcpynW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr, ByVal cchCount As Long) As LongPtr
#Else
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal pwsString As Long) As Long
Private Declare Function lstrcpynW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long, ByVal cchCount As Long) As Long
#End If


#If VBA7 Then
Function Utf8PtrToString(ByVal pUtf8String As LongPtr) As String
#Else
Function Utf8PtrToString(ByVal pUtf8String As Long) As String
#End If
    Dim Buffer As String
    Buffer = vbNullString
    Dim Length As Long
    Length = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)
    ' Length includes the terminating null character
    If Length > 1 Then
        Buffer = String(Length - 1, " ") ' and a termintating null char.
        Dim RetVal As Long
        RetVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, StrPtr(Buffer), Length)
        If RetVal = 0 Then
            Debug.Print "Utf8PtrToString Error:", Err.LastDllError
            Buffer = vbNullString
        End If
    End If
    Utf8PtrToString = Buffer
End Function


Function StringToUtf8Bytes(ByVal Text As String) As Variant
    Dim Length As Long
    Length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
    If Length = 0 Then
        Exit Function
    End If
    Dim Buffer() As Byte
    ReDim Buffer(Length)
    Dim RetVal As Long
    RetVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, VarPtr(Buffer(0)), Length, 0, 0)
    If RetVal = 0 Then
        Debug.Print "StringToUtf8Bytes Error:", Err.LastDllError
        Exit Function
    End If
    StringToUtf8Bytes = Buffer
End Function


#If VBA7 Then
Function Utf16PtrToString(ByVal pUtf16String As LongPtr) As String
#Else
Function Utf16PtrToString(ByVal pUtf16String As Long) As String
#End If
    Dim Length As Long
    Length = lstrlenW(pUtf16String)
    Utf16PtrToString = String(Length, " ")
    lstrcpynW StrPtr(Utf16PtrToString), pUtf16String, Length
End Function


''''===========================================================================
''''#If WIN64 Then
''''Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal pwsDest As LongPtr, ByVal pwsSource As LongPtr) As LongPtr
''''Private Declare PtrSafe Function SysAllocString Lib "OleAut32" (ByRef pwsString As LongPtr) As LongPtr
''''Private Declare PtrSafe Function SysStringLen Lib "OleAut32" (ByVal bstrString As LongPtr) As Long
''''#Else
''''Private Declare Function lstrcpyW Lib "kernel32" (ByVal pwsDest As Long, ByVal pwsSource As Long) As Long
''''Private Declare Function SysAllocString Lib "OleAut32" (ByRef pwsString As Long) As Long
''''Private Declare Function SysStringLen Lib "OleAut32" (ByVal bstrString As Long) As Long
''''#End If
