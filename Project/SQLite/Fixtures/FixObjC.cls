VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixObjC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides instantiated SQLiteC library objects"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides instantiated SQLiteC library objects"
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
''''
'''' Normally, dbm object is destructed before every request for new Connection
'''' object. dbm's destructor initiates reset of its descendants. However, the
'''' backup API test opens the source and destination databases on separate
'''' connections. When the source database is in-memory, it is destructed by the
'''' cleanup process initiated by requesting a second connection object for the
'''' destination database. In such a case, set .ReuseDBM to true to prevent
'''' cleanup. This attribute is automatically reset by the GetDBM routine.
''''
Option Explicit

Private Const LITE_LIB As String = "SQLiteCAdo"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TFixObjC
    DllMan As DllManager
    dbm As SQLiteC
    LibDir As String
    ReuseDBM As Boolean
End Type
Private this As TFixObjC


Private Sub Class_Initialize()
    this.LibDir = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX
End Sub

Friend Sub CleanUp()
    Set this.dbm = Nothing
    Set this.DllMan = Nothing
End Sub

Public Sub ReuseDBM()
    this.ReuseDBM = True
End Sub

Public Property Get DefaultDbPathName() As String
    DefaultDbPathName = this.LibDir & LITE_LIB & ".db"
End Property

Public Property Get DbMan() As SQLiteC
    Set DbMan = this.dbm
End Property

'@Description "Loads SQLite dll and returns DbManager (SQLiteC) object."
Public Function GetDBM() As SQLiteC
Attribute GetDBM.VB_Description = "Loads SQLite dll and returns DbManager (SQLiteC) object."
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    '''' Using default library names hardcoded in the SQLiteC constructor.
    If Not this.ReuseDBM Then
        Set this.dbm = Nothing
    Else
        this.ReuseDBM = False
    End If
    
    If this.dbm Is Nothing Then
        Set this.dbm = SQLiteC(DllPath)
        If this.dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
            "FixObjC", "Failed to create an SQLiteC instance."
    End If
    Set GetDBM = this.dbm
End Function

'@Description "Given database path, obtains a DbManager reference and returns a connection (SQLiteCConnection) object."
Public Function GetDBC(ByVal DbPathName As String, _
              Optional ByVal AllowNonExistent As Variant = True) As SQLiteCConnection
Attribute GetDBC.VB_Description = "Given database path, obtains a DbManager reference and returns a connection (SQLiteCConnection) object."
    Dim dbm As SQLiteC
    Set dbm = GetDBM()
    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(DbPathName, AllowNonExistent)
    If dbc Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create an SQLiteCConnection instance."
    Set GetDBC = dbc
End Function

'@Description "Returns a connection (SQLiteCConnection) object using the default project database."
Public Function GetDBCReg() As SQLiteCConnection
Attribute GetDBCReg.VB_Description = "Returns a connection (SQLiteCConnection) object using the default project database."
    Dim DbPathName As String
    DbPathName = this.LibDir & LITE_LIB & ".db"
    Set GetDBCReg = GetDBC(DbPathName)
End Function

'@Description "Returns a connection (SQLiteCConnection) object using a project database with Readonly file attribute."
Public Function GetDBCReadOnlyAttr() As SQLiteCConnection
Attribute GetDBCReadOnlyAttr.VB_Description = "Returns a connection (SQLiteCConnection) object using a project database with Readonly file attribute."
    Dim DbPathName As String
    DbPathName = this.LibDir & "ReadOnlyAttr.db"
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    With fso.GetFile(DbPathName)
        .Attributes = .Attributes Or ReadOnly
    End With
    Set GetDBCReadOnlyAttr = GetDBC(DbPathName)
End Function

'@Description "Returns a connection (SQLiteCConnection) object using anonymous database path."
Public Function GetDBCAnon() As SQLiteCConnection
Attribute GetDBCAnon.VB_Description = "Returns a connection (SQLiteCConnection) object using anonymous database path."
    Dim DbPathName As String
    DbPathName = ":blank:"
    Set GetDBCAnon = GetDBC(DbPathName)
End Function

'@Description "Returns a connection (SQLiteCConnection) object with invalid path."
Public Function GetDBCInvalidPath() As SQLiteCConnection
Attribute GetDBCInvalidPath.VB_Description = "Returns a connection (SQLiteCConnection) object with invalid path."
    Dim DbPathName As String
    DbPathName = "_:_/\_BAD PATH_<>;"
    Set GetDBCInvalidPath = GetDBC(DbPathName)
End Function

'@Description "Returns a connection (SQLiteCConnection) object using an in-memory database."
Public Function GetDBCMem(Optional ByVal OpenDb As Boolean = False) As SQLiteCConnection
Attribute GetDBCMem.VB_Description = "Returns a connection (SQLiteCConnection) object using an in-memory database."
    Dim AllowNonExistent As Boolean
    AllowNonExistent = True
    Dim DbPathName As String
    DbPathName = ":memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBC(DbPathName, AllowNonExistent)
    If OpenDb Then
        Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
            ErrNo.UnknownClassErr, "FixObjC", "Failed to open db."
    End If
    Set GetDBCMem = dbc
End Function

'@Description "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with an empty 'ITRB' table."
Public Function GetDBCMemITRB() As SQLiteCConnection
Attribute GetDBCMemITRB.VB_Description = "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with an empty 'ITRB' table."
    Set GetDBCMemITRB = GetDBCMem(OpenDb:=True) '''' DO NOT CLOSE In-Memory Database!
    GetDBCMemITRB.ExecuteNonQueryPlain FixSQLITRB.Create
End Function

'@Description "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with 'ITRB' table and data."
Public Function GetDBCMemITRBWithData() As SQLiteCConnection
Attribute GetDBCMemITRBWithData.VB_Description = "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with 'ITRB' table and data."
    Set GetDBCMemITRBWithData = GetDBCMem(OpenDb:=True)
    If CreateITRBTableWithData(GetDBCMemITRBWithData) <= 0 Then _
        Set GetDBCMemITRBWithData = Nothing
End Function

'''' Args:
''''   Blank (boolean, optional, false):
''''     If true, do not create data, that is act as GetDBCMem
''''
''''   UserVersion, AppId (long, optional, 0):
''''     If non-zero, set respective database attributes.
''''
'@Description "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with 'functions' table and data."
Public Function GetDBCMemFuncWithData(Optional ByVal Blank As Boolean = False, _
                                      Optional ByVal UserVersion As Long = 0, _
                                      Optional ByVal ApplicationId As Long = 0 _
                                      ) As SQLiteCConnection
Attribute GetDBCMemFuncWithData.VB_Description = "Returns a connection (SQLiteCConnection) object using created/opened in-memory database with 'functions' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCMem(OpenDb:=True)
    If Not Blank Then
        If CreateFunctionsTableWithData(dbc) <= 0 Then Set dbc = Nothing
    End If
    If (UserVersion Or ApplicationId) And Not dbc Is Nothing Then
        Dim dbs As SQLiteCStatement
        Set dbs = dbc.CreateStatement(vbNullString)
        If UserVersion <> 0 Then dbs.UserVersion = UserVersion
        If ApplicationId <> 0 Then dbs.ApplicationId = ApplicationId
    End If
    '''' DO NOT CLOSE In-Memory Database!
    Set GetDBCMemFuncWithData = dbc
End Function

'''' If database connection is not opened, it is opened and closed.
'@Description "Given SQLiteCConnection object, creates the 'functions' table with data."
Public Function CreateFunctionsTableWithData(ByVal dbc As SQLiteCConnection) As Long
Attribute CreateFunctionsTableWithData.VB_Description = "Given SQLiteCConnection object, creates the 'functions' table with data."
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.CreateWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    If HandleConnection Then Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
        ErrNo.UnknownClassErr, "FixObjC", "Failed to open db."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Guard.ExpressionTrueErr ResultCode <> SQLITE_OK, ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create table."
    Dim LogMsg As String
    LogMsg = "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    If HandleConnection Then Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, _
        ErrNo.UnknownClassErr, "FixObjC", "Failed to close db."
    CreateFunctionsTableWithData = AffectedRows
End Function

'''' If database connection is not opened, it is opened and closed.
'@Description "Given SQLiteCConnection object, creates the 'ITRB' table with data."
Public Function CreateITRBTableWithData(ByVal dbc As SQLiteCConnection) As Long
Attribute CreateITRBTableWithData.VB_Description = "Given SQLiteCConnection object, creates the 'ITRB' table with data."
    Dim HandleConnection As Boolean
    HandleConnection = (dbc.DbHandle = 0)
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    If HandleConnection Then Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, _
        ErrNo.UnknownClassErr, "FixObjC", "Failed to open db."
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Guard.ExpressionTrueErr ResultCode <> SQLITE_OK, ErrNo.UnknownClassErr, _
        "FixObjC", "Failed to create table."
    Dim LogMsg As String
    LogMsg = "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    Logger.Logg LogMsg, , DEBUGLEVEL_INFO
    If HandleConnection Then Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, _
        ErrNo.UnknownClassErr, "FixObjC", "Failed to close db."
    CreateITRBTableWithData = AffectedRows
End Function

'@Description "Generates a path pointing to a randomly named non-existing file in the temp folder."
Public Function RandomTempFileName(Optional ByVal FileExt As String = ".db") As String
Attribute RandomTempFileName.VB_Description = "Generates a path pointing to a randomly named non-existing file in the temp folder."
    RandomTempFileName = Environ$("TEMP") & PATH_SEP & _
        Format$(Now, "yyyy_mm_dd-hh_mm_ss") & "-" & Left$(GenerateGUID, 8) & FileExt
End Function

'@Description "Returns a connection (SQLiteCConnection) object using a path pointing to a newly created randomly named database in the temp folder."
Public Function GetDBCTmp(Optional ByVal NoInit As Boolean = False) As SQLiteCConnection
Attribute GetDBCTmp.VB_Description = "Returns a connection (SQLiteCConnection) object using a path pointing to a newly created randomly named database in the temp folder."
    Set GetDBCTmp = GetDBC(RandomTempFileName(), AllowNonExistent:=True)
    If Not NoInit Then
        '''' Make sure a blank db file is initialized
        With GetDBCTmp
            .OpenDb
            .Vacuum
            .CloseDb
        End With
    End If
End Function

'''' Args:
''''   Blank (boolean, optional, false):
''''     If true, do not create data, that is act as GetDBCTemp
''''
''''   UserVersion, AppId (long, optional, 0):
''''     If non-zero, set respective database attributes.
''''
'@Description "Returns a connection (SQLiteCConnection) object using a new random databse in the temp folder with 'functions' table and data."
Public Function GetDBCTmpFuncWithData(Optional ByVal Blank As Boolean = False, _
                                      Optional ByVal UserVersion As Long = 0, _
                                      Optional ByVal ApplicationId As Long = 0 _
                                      ) As SQLiteCConnection
Attribute GetDBCTmpFuncWithData.VB_Description = "Returns a connection (SQLiteCConnection) object using a new random databse in the temp folder with 'functions' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTmp()
    If Not Blank Then
        If CreateFunctionsTableWithData(dbc) <= 0 Then Set dbc = Nothing
    End If
    If UserVersion Or ApplicationId Then
        Guard.ExpressionTrueErr dbc.OpenDb <> SQLITE_OK, ErrNo.UnknownClassErr, _
            "FixObjC", "Failed to open db."
        Dim dbs As SQLiteCStatement
        Set dbs = dbc.CreateStatement(vbNullString)
        If UserVersion <> 0 Then dbs.UserVersion = UserVersion
        If ApplicationId <> 0 Then dbs.ApplicationId = ApplicationId
        Guard.ExpressionTrueErr dbc.CloseDb <> SQLITE_OK, ErrNo.UnknownClassErr, _
            "FixObjC", "Failed to close db."
    End If
    Set GetDBCTmpFuncWithData = dbc
End Function

'@Description "Returns a connection (SQLiteCConnection) object using a new random databse in the temp folder with 'ITRB' table and data."
Public Function GetDBCTmpITRBWithData() As SQLiteCConnection
Attribute GetDBCTmpITRBWithData.VB_Description = "Returns a connection (SQLiteCConnection) object using a new random databse in the temp folder with 'ITRB' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTmp()
    If CreateITRBTableWithData(dbc) <= 0 Then
        Set dbc = Nothing
        Set GetDBCTmpITRBWithData = Nothing
    Else
        Set GetDBCTmpITRBWithData = dbc
    End If
End Function

'@Description "Returns an SQLiteCStatement/ILiteADO object using created/opened in-memory database with 'ITRB' table and data."
Public Function GetDBSMemITRBWithData() As ILiteADO
Attribute GetDBSMemITRBWithData.VB_Description = "Returns an SQLiteCStatement/ILiteADO object using created/opened in-memory database with 'ITRB' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCMem(OpenDb:=True)
    If CreateITRBTableWithData(dbc) <= 0 Then
        Set dbc = Nothing
        Set GetDBSMemITRBWithData = Nothing
    Else
        Set GetDBSMemITRBWithData = dbc.CreateStatement(vbNullString)
    End If
End Function

'@Description "Returns an SQLiteCStatement/ILiteADO object using created/opened in-memory database with 'functions' table and data."
Public Function GetDBSMemFuncWithData() As ILiteADO
Attribute GetDBSMemFuncWithData.VB_Description = "Returns an SQLiteCStatement/ILiteADO object using created/opened in-memory database with 'functions' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCMem(OpenDb:=True) '''' DO NOT CLOSE In-Memory Database!
    If CreateFunctionsTableWithData(dbc) <= 0 Then
        Set dbc = Nothing
        Set GetDBSMemFuncWithData = Nothing
    Else
        Set GetDBSMemFuncWithData = dbc.CreateStatement(vbNullString)
    End If
End Function

'@Description "Returns an SQLiteCStatement/ILiteADO object using a new random databse in the temp folder with 'ITRB' table and data."
Public Function GetDBSTmpITRBWithData() As ILiteADO
Attribute GetDBSTmpITRBWithData.VB_Description = "Returns an SQLiteCStatement/ILiteADO object using a new random databse in the temp folder with 'ITRB' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTmp()
    If CreateITRBTableWithData(dbc) <= 0 Then
        Set dbc = Nothing
        Set GetDBSTmpITRBWithData = Nothing
    Else
        Set GetDBSTmpITRBWithData = dbc.CreateStatement(vbNullString)
    End If
End Function

'@Description "Returns an SQLiteCStatement/ILiteADO object using a new random databse in the temp folder with 'functions' table and data."
Public Function GetDBSTmpFuncWithData() As ILiteADO
Attribute GetDBSTmpFuncWithData.VB_Description = "Returns an SQLiteCStatement/ILiteADO object using a new random databse in the temp folder with 'functions' table and data."
    Dim dbc As SQLiteCConnection
    Set dbc = GetDBCTmp()
    If CreateFunctionsTableWithData(dbc) <= 0 Then
        Set dbc = Nothing
        Set GetDBSTmpFuncWithData = Nothing
    Else
        Set GetDBSTmpFuncWithData = dbc.CreateStatement(vbNullString)
    End If
End Function
