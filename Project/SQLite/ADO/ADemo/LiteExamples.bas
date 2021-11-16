Attribute VB_Name = "LiteExamples"
'@Folder "SQLite.ADO.ADemo"
'@IgnoreModule ProcedureNotUsed: This is a demo/examples module
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


'@Description "Checks and prints SQLiteODBC driver status."
Private Sub SQLite3ODBCDriverCheck()
Attribute SQLite3ODBCDriverCheck.VB_Description = "Checks and prints SQLiteODBC driver status."
    Debug.Print LiteMan.SQLite3ODBCDriverCheck()
End Sub

'@Description "Queries and prints SQLite version number."
Private Sub CheckConnectionAndVersion()
Attribute CheckConnectionAndVersion.VB_Description = "Queries and prints SQLite version number."
    Debug.Print LiteMan(":mem:").ExecADO.GetScalar(vbNullString)
End Sub

'@Description "Prints pathname of the created temp database."
Private Sub CreateTmpDb()
Attribute CreateTmpDb.VB_Description = "Prints pathname of the created temp database."
    Debug.Print LiteMan(":tmp:").ExecADO.MainDB
End Sub

'@Description "Clones default database and returns DB manager bound the new clone."
Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = FixObjAdo.DefaultDbName
    Dim DestinationDb As String
    DestinationDb = "Dest.db"
    Dim dbm As LiteMan
    Set dbm = LiteMan.CloneDb(DestinationDb, SourceDb)
    Debug.Print dbm.ExecADO.MainDB
End Sub
