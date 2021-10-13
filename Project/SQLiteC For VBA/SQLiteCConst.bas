Attribute VB_Name = "SQLiteCConst"
'@Folder "SQLiteC For VBA"
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

#If WIN64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
#End If

#If VBA7 <> True Then
    Public Const vbLongLong As Long = 20&
#End If

Public Const LITE_LIB As String = "SQLiteCforVBA"
Public Const PATH_SEP As String = "\"
Public Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP
