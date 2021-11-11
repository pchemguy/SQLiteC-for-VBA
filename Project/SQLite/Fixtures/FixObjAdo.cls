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

Private Const LITE_LIB As String = "SQLiteCAdo"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TFixObjADO
    LibDir As String
End Type
Private this As TFixObjADO


Private Sub Class_Initialize()
    this.LibDir = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX
End Sub

Public Property Get DefaultDbName() As String
    DefaultDbName = LITE_LIB & ".db"
End Property

Public Property Get DefaultDbPathName() As String
    DefaultDbPathName = this.LibDir & LITE_LIB & ".db"
End Property

Public Property Get DefaultDbPathNameRel() As String
    DefaultDbPathNameRel = LITE_RPREFIX & LITE_LIB & ".db"
End Property

Public Property Get RelPrefix() As String
    RelPrefix = LITE_RPREFIX
End Property

Public Function FixPath(Optional ByVal RelPath As String = vbNullString) As String
    FixPath = this.LibDir & "Fixtures" & PATH_SEP & RelPath
End Function

'@Description "Given database path returns a LiteADO/ILiteADO object suitable for working with the database."
Public Function GetDb(ByVal DbPathName As String) As ILiteADO
Attribute GetDb.VB_Description = "Given database path returns a LiteADO/ILiteADO object suitable for working with the database."
    Set GetDb = LiteMan(DbPathName).ExecADO
End Function

'@Description "Returns a LiteADO/ILiteADO object suitable for working the default database."
Public Function GetDbReg() As ILiteADO
Attribute GetDbReg.VB_Description = "Returns a LiteADO/ILiteADO object suitable for working the default database."
    Set GetDbReg = GetDb(this.LibDir & LITE_LIB & ".db")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an anonymous database."
Public Function GetDbAnon() As ILiteADO
Attribute GetDbAnon.VB_Description = "Returns a LiteADO/ILiteADO object for an anonymous database."
    Set GetDbAnon = GetDb(":blank:")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database."
Public Function GetDbMem() As ILiteADO
Attribute GetDbMem.VB_Description = "Returns a LiteADO/ILiteADO object for an in-memory database."
    Set GetDbMem = GetDb(":memory:")
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database with ITRB table defined."
Public Function GetDbMemITRB() As ILiteADO
Attribute GetDbMemITRB.VB_Description = "Returns a LiteADO/ILiteADO object for an in-memory database with ITRB table defined."
    Set GetDbMemITRB = GetDbMem()
    GetDbMemITRB.ExecuteNonQuery FixSQLITRB.Create
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database with ITRB table and data."
Public Function GetDbMemITRBWithData() As ILiteADO
Attribute GetDbMemITRBWithData.VB_Description = "Returns a LiteADO/ILiteADO object for an in-memory database with ITRB table and data."
    Set GetDbMemITRBWithData = GetDbMem()
    GetDbMemITRBWithData.ExecuteNonQuery FixSQLITRB.CreateWithData
End Function

'@Description "Returns a LiteADO/ILiteADO object for an in-memory database with Functions table and data."
Public Function GetDbMemFuncWithData() As ILiteADO
Attribute GetDbMemFuncWithData.VB_Description = "Returns a LiteADO/ILiteADO object for an in-memory database with Functions table and data."
    Set GetDbMemFuncWithData = GetDbMem()
    GetDbMemFuncWithData.ExecuteNonQuery FixSQLFunc.CreateWithData
End Function

'@Description "Generates a path pointing to a randomly named non-existing file in the temp folder."
Public Function RandomTempFileName(Optional ByVal FileExt As String = ".db") As String
Attribute RandomTempFileName.VB_Description = "Generates a path pointing to a randomly named non-existing file in the temp folder."
    RandomTempFileName = Environ$("TEMP") & PATH_SEP & _
        Format$(Now, "yyyy_mm_dd-hh_mm_ss") & "-" & Left$(GenerateGUID, 8) & FileExt
End Function

'@Description "Returns a LiteADO/ILiteADO object for a temp database."
Public Function GetDbTmp() As ILiteADO
Attribute GetDbTmp.VB_Description = "Returns a LiteADO/ILiteADO object for a temp database."
    Set GetDbTmp = GetDb(":temp:")
End Function

'@Description "Returns a LiteADO/ILiteADO object for a temp database."
Public Function GetDbTmpITRBWithData() As ILiteADO
Attribute GetDbTmpITRBWithData.VB_Description = "Returns a LiteADO/ILiteADO object for a temp database."
    Set GetDbTmpITRBWithData = GetDb(":temp:")
    GetDbTmpITRBWithData.ExecuteNonQuery FixSQLITRB.CreateWithData
End Function

'@Description "Returns a LiteADO/ILiteADO object for a temp database."
Public Function GetDbTmpFuncWithData() As ILiteADO
Attribute GetDbTmpFuncWithData.VB_Description = "Returns a LiteADO/ILiteADO object for a temp database."
    Set GetDbTmpFuncWithData = GetDb(":temp:")
    GetDbTmpFuncWithData.ExecuteNonQuery FixSQLFunc.CreateWithData
End Function
