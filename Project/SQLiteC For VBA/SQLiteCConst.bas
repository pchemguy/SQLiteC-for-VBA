Attribute VB_Name = "SQLiteCConst"
'@Folder "SQLiteC For VBA"
'@IgnoreModule IndexedDefaultMemberAccess

''''======================================================================''''
'''' Acknowledgement
'''' Some code from the https://github.com/govert/SQLiteForExcel project.
''''======================================================================''''

Option Explicit

#If WIN64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
#End If

#If VBA7 <> True Then
    Public Const vbLongLong As Long = 20&
#End If

Public Const KeyAlreadyExistsErr As Long = 457
Public Const OutOfMemoryErr As Long = 7&
Public Const ConnectionNotOpenedErr As Long = vbObjectError + 3000
Public Const StatementNotPreparedErr As Long = vbObjectError + 3001
