Attribute VB_Name = "TableOExamples"
'@Folder("SQLiteDB.DB Objects")
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


'@Description "Sets sample TableO workflow - populate from existing db table metadata, modify, generate SQL create code."
'@EntryPoint
Public Sub TableODevHelper()
Attribute TableODevHelper.VB_Description = "Sets sample TableO workflow - populate from existing db table metadata, modify, generate SQL create code."
    Dim Database As String
    
    Dim ExtPattern As String
    ExtPattern = "\.\w+$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = ExtPattern
    Database = re.Replace(ThisWorkbook.Name, ".db")
    
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(Database)
    
    Dim TableName As String
    TableName = "test_table"
    
    Dim DbTable As TableO
    Set DbTable = TableO.FromDb(TableName, DbManager)
    
    Dim DbFields As Scripting.Dictionary
    Set DbFields = DbTable.Fields
    
    Dim FieldIndex As Long
    For FieldIndex = LBound(DbTable.FieldNames) To UBound(DbTable.FieldNames)
        Debug.Print DbFields(DbTable.FieldNames(FieldIndex)).SQL
    Next FieldIndex
        
    Dim Cons As Scripting.Dictionary
    Dim ConName As Variant
    
    Dim ConCK As ConstraintCK
    Set Cons = DbTable.CKs
    If Cons.Count > 0 Then
        For Each ConName In Cons.Keys
            Set ConCK = Cons(ConName)
            Debug.Print ConCK.SQL
        Next ConName
    End If
    
    Dim ConUQ As ConstraintUQ
    Set Cons = DbTable.UQs
    If Cons.Count > 0 Then
        For Each ConName In Cons.Keys
            Set ConUQ = Cons(ConName)
            Debug.Print ConUQ.SQL
        Next ConName
    End If
    
    Dim ConFK As ConstraintFK
    Set Cons = DbTable.FKs
    If Cons.Count > 0 Then
        For Each ConName In Cons.Keys
            Set ConFK = Cons(ConName)
            Debug.Print ConFK.SQL
        Next ConName
    End If
    
    If Not DbTable.PK Is Nothing Then
        Dim ConPK As ConstraintPK
        Set ConPK = DbTable.PK
        Debug.Print ConPK.SQL
    End If
    
    Debug.Print "=================================================="
    DbTable.CreateActionIfExists = vbNullString
    Debug.Print DbTable.SQL
    Debug.Print "=================================================="
    Debug.Print "=================================================="
    DbTable.CreateActionIfExists = "DROP"
    Debug.Print DbTable.SQL
    Debug.Print "=================================================="
    Debug.Print "=================================================="
    DbTable.CreateActionIfExists = "SKIP"
    Debug.Print DbTable.SQL
    Debug.Print "=================================================="
    
    DbTable.TableName = "__bebe__"
    DbTable.CreateTable "SKIP"
End Sub
