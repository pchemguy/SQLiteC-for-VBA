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
'@Folder "SQLite.C.UTF Library"
'@ModuleDescription "Performs UTF text conversion"
'@PredeclaredId
Option Explicit

'@Ignore UseMeaningfulName
Private Const CP_UTF8 As Long = 65001

#If Win64 Then
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
Public Function StrFromUTF8Ptr(ByVal UTF8Ptr As LongPtr) As String
#Else
Public Function StrFromUTF8Ptr(ByVal UTF8Ptr As Long) As String
#End If
    If UTF8Ptr = 0 Then
        StrFromUTF8Ptr = vbNullString
        Exit Function
    End If
    Dim Buffer As String
    Dim Length As Long
    Length = MultiByteToWideChar(CP_UTF8, 0, UTF8Ptr, -1, 0, 0)
    ' Length includes the terminating null character
    If Length > 1 Then
        Buffer = String(Length - 1, " ") ' and a termintating null char.
        Dim RetVal As Long
        RetVal = MultiByteToWideChar(CP_UTF8, 0, UTF8Ptr, -1, StrPtr(Buffer), Length)
        If RetVal = 0 Then
            Debug.Print "StrFromUTF8Ptr Error:", Err.LastDllError
            Buffer = vbNullString
        End If
    End If
    StrFromUTF8Ptr = Buffer
End Function


Public Function UTF8BytesFromStr(ByVal U16Str As String) As Variant
    Dim Length As Long
    Length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(U16Str), -1, 0, 0, 0, 0)
    If Length = 0 Then
        Exit Function
    End If
    Dim Buffer() As Byte
    ReDim Buffer(Length)
    Dim RetVal As Long
    RetVal = WideCharToMultiByte(CP_UTF8, 0, StrPtr(U16Str), -1, VarPtr(Buffer(0)), Length, 0, 0)
    If RetVal = 0 Then
        Debug.Print "UTF8BytesFromStr Error:", Err.LastDllError
        Exit Function
    End If
    UTF8BytesFromStr = Buffer
End Function


'@Ignore ProcedureNotUsed
#If VBA7 Then
Public Function StrFromUTF16Ptr(ByVal UTF16Ptr As LongPtr) As String
#Else
Public Function StrFromUTF16Ptr(ByVal UTF16Ptr As Long) As String
#End If
    If UTF16Ptr = 0 Then
        StrFromUTF16Ptr = vbNullString
        Exit Function
    End If
    Dim Length As Long
    Length = lstrlenW(UTF16Ptr)
    StrFromUTF16Ptr = String(Length, " ")
    lstrcpynW StrPtr(StrFromUTF16Ptr), UTF16Ptr, Length
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
