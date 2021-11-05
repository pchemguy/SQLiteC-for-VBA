VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixObjAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides instantiated SQLiteADO library objects"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides instantiated SQLiteADO library objects"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Private Const LITE_LIB As String = "SQLiteCDBVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TFixObjADO
    LibDir As String
End Type
Private this As TFixObjADO


Private Sub Class_Initialize()
    this.LibDir = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX
End Sub

Public Property Get DefaultDbPathName() As String
    DefaultDbPathName = this.LibDir & LITE_LIB & ".db"
End Property

'@Description "Given database path returns a LiteADO/ILiteADO object suitable for working with the database."
Public Function GetDBM(ByVal DbPathName As String) As ILiteADO
    Set GetDBM = LiteADO.Create(DbPathName)
End Function

'@Description "Returns a LiteADO/ILiteADO object suitable for working the default database."
Public Function GetDBMReg() As ILiteADO
    Set GetDBMReg = GetDBM(this.LibDir & LITE_LIB & ".db")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an anonymous database."
Public Function GetDBMAnon() As ILiteADO
    Set GetDBMAnon = GetDBM(":blank:")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database."
Public Function GetDBMMem() As ILiteADO
    Set GetDBMMem = GetDBM(":memory:")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database with ITRB table defined."
Public Function GetDBMMemITRB() As ILiteADO
    Dim dbm As ILiteADO
    Set dbm = GetDBM(":memory:")
    Dim AffectedRecords As Long
    AffectedRecords = dbm.ExecuteNonQuery(FixSQLITRB.Create)
    Debug.Assert AffectedRecords = 0
    Set GetDBMMemITRB = dbm
End Function

'@Description "Generates a path pointing to a randomly named non-existing file in the temp folder."
Public Function RandomTempFileName(Optional ByVal FileExt As String = ".db") As String
Attribute RandomTempFileName.VB_Description = "Generates a path pointing to a randomly named non-existing file in the temp folder."
    RandomTempFileName = Environ$("TEMP") & PATH_SEP & _
        Format$(Now, "yyyy_mm_dd-hh_mm_ss") & "-" & Left$(GenerateGUID, 8) & FileExt
End Function
