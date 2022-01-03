VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Wraps SQLite C keywords API."
'@Folder "SQLite.SQLQuoter"
'@ModuleDescription "Wraps SQLite C keywords API."
'@PredeclaredId
Option Explicit

Private Const SQLITE_N_KEYWORD As Long = 147

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)
#End If

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_keyword_count Lib "SQLite3" () As Long
Private Declare PtrSafe Function sqlite3_keyword_name Lib "SQLite3" (ByVal kwi As Long, ByRef pzName As LongPtr, ByRef pnName As Long) As Long
#Else
Private Declare Function sqlite3_keyword_count Lib "SQLite3" () As Long
Private Declare Function sqlite3_keyword_name Lib "SQLite3" (ByVal kwi As Long, ByRef pzName As Long, ByRef pnName As Long) As Long
#End If

Private Type TObjectState
    DllMan As DllManager
    SQLiteKeywords As Scripting.Dictionary
End Type
'@Ignore MoveFieldCloserToUsage
Private this As TObjectState


Public Function Create(Optional ByVal DllMan As DllManager = Nothing _
                      ) As SQLiteCKeyword
    Dim Instance As SQLiteCKeyword
    Set Instance = New SQLiteCKeyword
    Instance.Init DllMan
    Set Create = Instance
End Function

Friend Sub Init(Optional ByVal DllMan As DllManager = Nothing)
    Set this.SQLiteKeywords = New Scripting.Dictionary
    this.SQLiteKeywords.CompareMode = TextCompare
    If Not DllMan Is Nothing Then
        Set this.DllMan = DllMan
    Else
        Dim FileNames As Variant
        #If Win64 Then
            FileNames = "sqlite3.dll"
        #Else
            FileNames = Array( _
                "icudt" & SQL_ICU_V & ".dll", "icuuc" & SQL_ICU_V & ".dll", _
                "icuin" & SQL_ICU_V & ".dll", "icuio" & SQL_ICU_V & ".dll", _
                "icutu" & SQL_ICU_V & ".dll", "sqlite3.dll")
        #End If
        Set this.DllMan = DllManager.Create(vbNullString, FileNames)
    End If
    If KeywordCount() > SQLITE_N_KEYWORD Then
        Dim KeywordIndex As Long
        KeywordIndex = 0
        Dim Keyword As String
        Keyword = NthKeyword(KeywordIndex)
        Do While Keyword <> vbNullString
            '@Ignore IndexedDefaultMemberAccess
            this.SQLiteKeywords(Keyword) = KeywordIndex
            KeywordIndex = KeywordIndex + 1
            Keyword = NthKeyword(KeywordIndex)
        Loop
    End If
End Sub

Public Property Get SQLiteKeywords() As Scripting.Dictionary
    Set SQLiteKeywords = this.SQLiteKeywords
End Property

'''' Reference: https://www.sqlite.org/c3ref/keyword_check.html
''''
Friend Function KeywordCount() As Long
    KeywordCount = sqlite3_keyword_count()
End Function

'''' Reference: https://www.sqlite.org/c3ref/keyword_check.html
''''
Friend Function NthKeyword(ByVal KeywordIndex As Long) As String
    #If VBA7 Then
        Dim KeywordPtr As LongPtr
    #Else
        Dim KeywordPtr As Long
    #End If
    KeywordPtr = 0
    Dim KeywordLength As Long
    KeywordLength = 0
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_keyword_name(KeywordIndex, KeywordPtr, KeywordLength)
    If ResultCode <> SQLITE_OK Or KeywordPtr = 0 Or KeywordLength = 0 Then
        NthKeyword = vbNullString
    Else
        Dim KeywordBuffer() As Byte
        ReDim KeywordBuffer(0 To KeywordLength - 1)
        CopyMemory KeywordBuffer(0), KeywordPtr, KeywordLength
        NthKeyword = StrConv(KeywordBuffer, vbUnicode)
    End If
End Function


''''    Private Declare PtrSafe Function sqlite3_keyword_check Lib "SQLite3" (ByVal zName As LongPtr, ByVal nName As Long) As Long
''''    Private Declare Function sqlite3_keyword_check Lib "SQLite3" (ByVal zName As Long, ByVal nName As Long) As Long
