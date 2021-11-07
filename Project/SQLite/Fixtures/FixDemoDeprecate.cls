VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixDemoDeprecate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides fixtures for testing and as usage patterns"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides fixtures for testing and as usage patterns"
'@PredeclaredId
Option Explicit

Private Const LITE_LIB As String = "SQLiteCAdo"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

'''' Custom functions added to SQLite source for testing/verification purposes
#If VBA7 Then
    #If WIN64 Then
    Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As LongLong
    #Else
    Private Declare PtrSafe Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
    #End If
Private Declare PtrSafe Function sqlite3_latin_utf8 Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_cyrillic_utf8 Lib "SQLite3" () As LongPtr ' PtrUtf8String
#Else
Private Declare Function sqlite3_libversion_number_i64 Lib "SQLite3" () As Currency
Private Declare Function sqlite3_latin_utf8 Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_cyrillic_utf8 Lib "SQLite3" () As Long ' PtrUtf8String
#End If

Private Type TSQLiteCConnDemoFix
    dbm As SQLiteC
    DbConn As SQLiteCConnection
    DbStmt As SQLiteCStatement
    FixPath As String
End Type
Private this As TSQLiteCConnDemoFix


Public Function Create() As FixDemoDeprecate
    Dim Instance As FixDemoDeprecate
    Set Instance = New FixDemoDeprecate
    Instance.Init
    Set Create = Instance
End Function

Friend Sub Init()
    Dim DllPath As String
    DllPath = "Library\SQLiteCforVBA\dll\" & ARCH
    '@Ignore IndexedDefaultMemberAccess
    Set this.dbm = SQLiteC(DllPath)
    this.FixPath = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & "Fixtures" & PATH_SEP
End Sub

Public Property Get dbm() As SQLiteC
    Set dbm = this.dbm
End Property

Private Property Get ConnDb(ByVal DbPathName As String) As SQLiteCConnection
    If Not this.DbConn Is Nothing Then
        Set ConnDb = this.DbConn
    Else
        Dim dbm As SQLiteC
        Set dbm = this.dbm
        Dim DbConn As SQLiteCConnection
        Set DbConn = dbm.CreateConnection(DbPathName)
        Set this.DbConn = DbConn
        Set ConnDb = DbConn
    End If
End Property

Public Property Get StmtDb(Optional ByVal StmtName As String = vbNullString) As SQLiteCStatement
    Set StmtDb = this.DbConn.CreateStatement(StmtName)
End Property

Public Property Get ConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set ConnDbRegular = ConnDb(DbPathName)
End Property

'@Ignore ProcedureNotUsed
Public Property Get ConnDbMemory() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ":memory:"
    Set ConnDbMemory = ConnDb(DbPathName)
End Property

'@Ignore ProcedureNotUsed
Public Property Get ConnDbTemp() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = vbNullString
    Set ConnDbTemp = ConnDb(DbPathName)
End Property

Public Property Get ConnDbInvalidPath() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set ConnDbInvalidPath = ConnDb(DbPathName)
End Property

Public Property Get ConnDbNotDb() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.FixPath & "NotDb.db"
    Set ConnDbNotDb = ConnDb(DbPathName)
End Property

Public Property Get ConnDbLockedDb() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = this.FixPath & "ACLLocked" & PATH_SEP & "LockedDb.db"
    Set ConnDbLockedDb = ConnDb(DbPathName)
End Property

Public Function LibVersionNumber() As Variant
    LibVersionNumber = sqlite3_libversion_number_i64()
    If VarType(LibVersionNumber) = vbCurrency Then
        LibVersionNumber = LibVersionNumber * 10000
    End If
End Function


'''' =========================================================================== ''''
'''' This routine uses a non-standard export added to a the SQLite amalgamation  ''''
'''' yielding a custom-built library. It returns SQLite version as a 64-bit int. ''''
'''' It can be used for testing purposes, especially for the x32 version.        ''''
'''' =========================================================================== ''''


Public Function LatinUTF8() As String
    LatinUTF8 = UTFlib.StrFromUTF8Ptr(sqlite3_latin_utf8())
End Function

Public Function CyrillicUTF8() As String
    CyrillicUTF8 = UTFlib.StrFromUTF8Ptr(sqlite3_cyrillic_utf8())
End Function

'@Ignore UseMeaningfulName
Public Function VersionI64() As Variant
    Dim VerNum As Variant
    VerNum = sqlite3_libversion_number_i64()
    If VarType(VerNum) = vbCurrency Then VerNum = VerNum * 10000
    VersionI64 = VerNum
End Function
