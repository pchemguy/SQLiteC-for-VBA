VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteMetaSQLIdxFK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "SQL queries for retrieving detailed information on database indices and foreign keys."
'@Folder "SQLiteDB.Introspection"
'@ModuleDescription "SQL queries for retrieving detailed information on database indices and foreign keys."
'@PredeclaredId
''''
'''' Logically, this module is a part of SQLiteSQLDbInfo, and this FK/IDX code is
'''' placed in a separate module simply to isolate the large amount of SQL code.
'''' All methods of this module are exposed by SQLiteSQLDbInfo via composition.
'''' This class is not supposed to be used directly, and it does not need to be
'''' instantiated: all functionality can be used via the default instance.
''''
Option Explicit


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning database tables, skipping
'''' system tables (prefixed with "sqlite_") and ordering by ROWID
'''' (in order of creation). If requested, a CTE WITH term is
'''' generated.
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''   CTEWITH (boolean, optional, False):
''''     If True, format as a CTE WITH term
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.Tables
''''   SELECT name, sql
''''   FROM main.sqlite_master
''''   WHERE type = 'table' AND (name NOT LIKE 'sqlite_%')
''''   ORDER BY ROWID ASC
''''
''''   >>> ?LiteMetaSQLIdxFK.Tables(, True)
''''   t AS (
''''       SELECT name, sql
''''       FROM main.sqlite_master
''''       WHERE type = 'table' AND (name NOT LIKE 'sqlite_%')
''''       ORDER BY ROWID ASC
''''   )
''''
'@Description "Generates a query returning database tables."
Public Function Tables(Optional ByVal Schema As String = "main", _
                       Optional ByVal CTEWITH As Boolean = False) As String
Attribute Tables.VB_Description = "Generates a query returning database tables."
    Dim Indent As String
    Dim Query As String
    Indent = IIf(CTEWITH, "    ", vbNullString)
    Query = Indent & Join(Array( _
        "SELECT tbl_name, sql", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE type = 'table' AND (name NOT LIKE 'sqlite_%')", _
        "ORDER BY ROWID ASC" _
    ), vbNewLine & Indent)
    Tables = IIf(CTEWITH, "t AS (" & vbNewLine & Query & vbNewLine & ")", Query)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning database views ordered by ROWID
'''' (in order of creation).
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.Views
''''   SELECT tbl_name, sql
''''   FROM main.sqlite_master
''''   WHERE type = 'view'
''''   ORDER BY ROWID ASC
''''
'@Ignore ProcedureNotUsed
'@Description "Generates a query returning database views."
Public Function Views(Optional ByVal Schema As String = "main") As String
Attribute Views.VB_Description = "Generates a query returning database views."
    Views = Join(Array( _
        "SELECT tbl_name, sql", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE type = 'view'", _
        "ORDER BY ROWID ASC" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning database triggers ordered by ROWID
'''' (in order of creation).
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.Triggers
''''   SELECT tbl_name, sql
''''   FROM main.sqlite_master
''''   WHERE type = 'trigger'
''''   ORDER BY ROWID ASC
''''
'@Description "Generates a query returning database triggers."
Public Function Triggers(Optional ByVal Schema As String = "main") As String
Attribute Triggers.VB_Description = "Generates a query returning database triggers."
    Triggers = Join(Array( _
        "SELECT tbl_name, sql", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE type = 'trigger'", _
        "ORDER BY ROWID ASC" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning all non-system database objects
'''' ordered by type (tables, indices, views, triggers) and then by ROWID.
'''' The query returns two columns (sql, type_id).
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.DbSchema
''''   SELECT sql, (CASE type
''''       WHEN 'table' THEN 0
''''       WHEN 'index' THEN 1
''''       WHEN 'view'  THEN 3
''''                    ELSE 4
''''                    END) AS type_id
''''   FROM main.sqlite_master
''''   WHERE name NOT like 'sqlite_%'
''''   ORDER BY type_id, _ROWID_
''''
'@Description "Generates a query returning all non-system database objects."
Public Function DbSchema(Optional ByVal Schema As String = "main") As String
Attribute DbSchema.VB_Description = "Generates a query returning all non-system database objects."
    DbSchema = Join(Array( _
        "SELECT sql, (CASE type", _
        "    WHEN 'table' THEN 0", _
        "    WHEN 'index' THEN 1", _
        "    WHEN 'view'  THEN 2", _
        "                 ELSE 3", _
        "                 END) AS type_id", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE name NOT like 'sqlite_%'", _
        "ORDER BY type_id, _ROWID_" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning all non-system database objects,
'''' except for triggers, ordered by type (tables, indices, views) and
'''' then by ROWID. The query returns two columns (sql, type_id).
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.DbSchemaNoTriggers
''''   SELECT sql, (CASE type
''''       WHEN 'table' THEN 0
''''       WHEN 'index' THEN 1
''''                    ELSE 2
''''                    END) AS type_id
''''   FROM main.sqlite_master
''''   WHERE (name NOT like 'sqlite_%') AND type <> 'trigger'
''''   ORDER BY type_id, _ROWID_
''''
'@Description "Generates a query returning all non-system database objects."
Public Function DbSchemaNoTriggers(Optional ByVal Schema As String = "main") As String
Attribute DbSchemaNoTriggers.VB_Description = "Generates a query returning all non-system database objects."
    DbSchemaNoTriggers = Join(Array( _
        "SELECT sql, (CASE type", _
        "    WHEN 'table' THEN 0", _
        "    WHEN 'index' THEN 1", _
        "                 ELSE 2", _
        "                 END) AS type_id", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE (name NOT like 'sqlite_%') AND type <> 'trigger'", _
        "ORDER BY type_id, _ROWID_" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning base info on database indices ordering
'''' by ROWID (in order of creation). If requested, a CTE WITH term is generated.
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''   CTEWITH (boolean, optional, False):
''''     If True, format as a CTE WITH term
''''
'''' Returns:
''''   String, containing the query
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.IndexBase
''''   SELECT ROWID AS id, name AS idx_name, tbl_name, sql
''''   FROM main.sqlite_master
''''   WHERE type='index'
''''   ORDER BY ROWID ASC
''''
''''   >>> ?LiteMetaSQLIdxFK.IndexBase(, True)
''''   ib AS (
''''       SELECT ROWID AS id, name AS idx_name, tbl_name, sql
''''       FROM main.sqlite_master
''''       WHERE type='index'
''''       ORDER BY ROWID ASC
''''   )
''''
'@Description "Generates a query returning indices (base info)."
Public Function IndexBase(Optional ByVal Schema As String = "main", _
                          Optional ByVal CTEWITH As Boolean = False) As String
Attribute IndexBase.VB_Description = "Generates a query returning indices (base info)."
    Dim Indent As String
    Dim Query As String
    Indent = IIf(CTEWITH, "    ", vbNullString)
    Query = Indent & Join(Array( _
        "SELECT ROWID AS id, name AS idx_name, tbl_name, sql", _
        "FROM " & Schema & ".sqlite_master", _
        "WHERE type = 'index'", _
        "ORDER BY ROWID ASC" _
    ), vbNewLine & Indent)
    IndexBase = IIf(CTEWITH, "ib AS (" & vbNewLine & Query & vbNewLine & ")", Query)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite CTE WITH term for a foreign key list.
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the CTE WITH term
''''
'''' Examples:
''''   >>> ?LiteMetaSQLIdxFK.pForeignKeyList
''''   fkl AS (
''''       SELECT tbl_name AS child_table, [from] AS child_col0,
''''              [table] AS parent_table, [to] AS parent_col0,
''''              on_update, on_delete, id AS fk_id, seq AS fk_seq
''''       FROM t
''''       Join main.pragma_foreign_key_list(t.tbl_name)
''''       ORDER BY child_table, fk_id
''''   )
''''
'@Description "Generates a query returning a foreign key CTE WITH term."
Public Function pForeignKeyList(Optional ByVal Schema As String = "main") As String
Attribute pForeignKeyList.VB_Description = "Generates a query returning a foreign key CTE WITH term."
    pForeignKeyList = Join(Array( _
        "fkl AS (", _
        "    SELECT tbl_name AS child_table, [from] AS child_col0,", _
        "           [table] AS parent_table, [to] AS parent_col0,", _
        "           on_update, on_delete, id AS fk_id, seq AS fk_seq", _
        "    FROM t", _
        "    JOIN " & Schema & ".pragma_foreign_key_list(t.tbl_name)", _
        "    ORDER BY child_table, fk_id", _
        "),", _
        "fk AS (", _
        "    SELECT *, group_concat(child_col0, ', ') AS child_cols,", _
        "              group_concat(parent_col0, ', ') AS parent_cols,", _
        "              min(fk_seq) AS min_fk_seq", _
        "    FROM fkl", _
        "    GROUP BY child_table, fk_id", _
        "    ORDER BY child_table, fk_id", _
        ")" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite CTE WITH term for index info & list.
'''' For each index list info and join the tables. Only use <index name> here.
'''' For multi-column indices, keep the row with the first column and generates
'''' a column list. Generate database-wide list of additional index info columns
'''' from the per-table index lists.
''''
'''' Args:
''''   Schema (string, optional, "main"):
''''     Schema name/alias
''''
'''' Returns:
''''   String, containing the CTE WITH term
''''
'@Description "Generates a query returning a CTE WITH term for index info & list."
Public Function pIndexInfoList(Optional ByVal Schema As String = "main") As String
Attribute pIndexInfoList.VB_Description = "Generates a query returning a CTE WITH term for index info & list."
    pIndexInfoList = Join(Array( _
        "ii AS (", _
        "    SELECT ib.idx_name, min(ii.seqno) AS seqno, ii.name AS col0_name, group_concat(ii.name, ', ') AS columns", _
        "    FROM ib", _
        "    JOIN " & Schema & ".pragma_index_info(ib.idx_name) AS ii", _
        "    GROUP BY idx_name", _
        "),", _
        "il AS (", _
        "    SELECT name AS idx_name, seq AS idx_seq, [unique], origin, partial", _
        "    FROM t", _
        "    JOIN " & Schema & ".pragma_index_list(tbl_name)", _
        ")" _
    ), vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'@Description "Generates a query returning all foreing keys in the SQLite database"
Public Function ForeingKeys(Optional ByVal Schema As String = "main") As String
Attribute ForeingKeys.VB_Description = "Generates a query returning all foreing keys in the SQLite database"
    Dim StmtParts(0 To 5) As String
    StmtParts(0) = "WITH"
    
    '''' List all db tables
    StmtParts(1) = Tables(Schema, True) & ","
    
    '''' For each table, list foreign keys and join them to get a list of all foreign
    '''' keys for the DB. Each row contains info on a foreign key for a single column.
    '''' Yield a single row per foreign key, including multi-column keys. For multi-column
    '''' keys, keep the row with the first column and generates a column list.
    StmtParts(2) = pForeignKeyList(Schema)
    
    StmtParts(3) = "SELECT *"
    StmtParts(4) = "FROM fk AS foreign_keys"
    StmtParts(5) = "ORDER BY child_table, fk_id"
    
    ForeingKeys = Join(StmtParts, vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates an SQLite query returning database indices, ordering by ROWID.
'''' If "NonSys" = True, skip auto indices (prefixed with "sqlite_autoindex_").
''''
'@Description "Generates a query returning all indices in the SQLite database"
Public Function Indices(Optional ByVal Schema As String = "main", _
                        Optional ByVal NonSys As Boolean = True) As String
Attribute Indices.VB_Description = "Generates a query returning all indices in the SQLite database"
    Dim StmtParts(10 To 26) As String
    StmtParts(10) = "WITH"
    
    '''' List all db tables
    StmtParts(11) = Tables(Schema, True) & ","
    
    '''' List all db indices
    StmtParts(12) = IndexBase(Schema, True) & ","
    
    '''' For each index list info and join the tables. Only use <index name> here. For
    '''' multi-column indices, keep the row with the first column and generates a column list.
    '''' Generate database-wide list of additional index info columns from the per-table index lists
    StmtParts(13) = pIndexInfoList(Schema) & ","
    
    '''' After taking care of multi-row descriptions, add aditional columns from index list
    StmtParts(14) = "idx AS ("
    StmtParts(15) = "    SELECT ib.id, ib.idx_name, ib.tbl_name, ii.col0_name, ii.columns, ib.sql"
    StmtParts(16) = "    FROM ib, ii"
    StmtParts(17) = "    ON ib.idx_name = ii.idx_name"
    StmtParts(18) = "),"
    
    '''' Join additional info columns with index-wise list
    StmtParts(19) = "iex AS ("
    StmtParts(20) = "    SELECT idx.*, il.idx_seq, il.[unique], il.origin, il.partial"
    StmtParts(21) = "    FROM idx, il"
    StmtParts(22) = "    WHERE idx.idx_name = il.idx_name"
    StmtParts(23) = ")"
    
    StmtParts(24) = "SELECT *"
    StmtParts(25) = "FROM iex AS indices"
    StmtParts(26) = IIf(NonSys, _
                    "WHERE idx_name NOT LIKE 'sqlite_autoindex_%'" & vbNewLine, vbNullString) & _
                    "ORDER BY id"
    
    Indices = Join(StmtParts, vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Indices on child columns of foreing key relations are not mandatory,
'''' but generally should be defined. Database engine does not control whether
'''' such indices are defined. This query return a summary table showing all
'''' child columns and corresponding indices in the "idx_name" column. If this
'''' field is empty for a particular child column, the corresponding index has
'''' not been defined.
''''
'@Description "Generates a query returning child columns for all foreing keys and corresponding indices."
Public Function FKChildIndices(Optional ByVal Schema As String = "main") As String
Attribute FKChildIndices.VB_Description = "Generates a query returning child columns for all foreing keys and corresponding indices."
    Dim StmtParts(10 To 34) As String
    StmtParts(10) = "WITH"
    StmtParts(11) = Tables(Schema, True) & ","
    StmtParts(12) = IndexBase(Schema, True) & ","
    StmtParts(13) = pIndexInfoList(Schema) & ","
    StmtParts(14) = "idx AS ("
    StmtParts(15) = "    SELECT ib.id, ib.idx_name, ib.tbl_name, ii.col0_name, ii.columns, ib.sql"
    StmtParts(16) = "    FROM ib, ii"
    StmtParts(17) = "    ON ib.idx_name = ii.idx_name"
    StmtParts(18) = "),"
    StmtParts(19) = "iex AS ("
    StmtParts(20) = "    SELECT idx.*, il.idx_seq, il.[unique], il.origin, il.partial"
    StmtParts(21) = "    FROM idx, il"
    StmtParts(22) = "    WHERE idx.idx_name = il.idx_name AND partial = 0"
    StmtParts(23) = "),"

    StmtParts(24) = pForeignKeyList(Schema) & ","

    '''' Join indices and foreign keys tables to see which child columns do not have indices.
    '''' Multi-column indices, having the child column set as the "prefix" are accepted.
    StmtParts(25) = "fki AS ("
    StmtParts(26) = "    SELECT fk.child_table, fk.child_cols, fk.parent_table, fk.parent_cols,"
    StmtParts(27) = "           iex.idx_name"
    StmtParts(28) = "    FROM fk"
    StmtParts(29) = "    LEFT JOIN iex"
    StmtParts(30) = "    ON fk.child_table = iex.tbl_name AND fk.child_cols = substr(iex.columns, 1, length(fk.child_cols))"
    StmtParts(31) = ")"
    StmtParts(32) = "SELECT *"
    StmtParts(33) = "FROM fki AS fkeys_childindices"
    StmtParts(34) = "ORDER BY child_table, child_cols"

    FKChildIndices = Join(StmtParts, vbNewLine)
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' If IDX1 indexes columns (A, B) and IDX2 indexes columns (A, B, C), that is
'''' IDX1 indexes a "prefix" of IDX2, IDX2 can replace IDX1. On the other hand,
'''' depending on statistics (if for any given pair (A, B), there are very few
'''' rows), IDX2 may not be justifiable (unless it is the primary key). This
'''' query aims to return all such similar ("prefix") indices, though it has not
'''' been thoughroughly verified. It may return some "false" positive. Whether
'''' it can miss indices is not clear.
''''
'@Description "Generates a query returning similar indices."
Public Function SimilarIndices(Optional ByVal Schema As String = "main") As String
Attribute SimilarIndices.VB_Description = "Generates a query returning similar indices."
    Dim StmtParts(10 To 39) As String
    StmtParts(10) = "WITH"
    StmtParts(11) = Tables(Schema, True) & ","
    StmtParts(12) = IndexBase(Schema, True) & ","
    StmtParts(13) = pIndexInfoList(Schema) & ","
    StmtParts(14) = "idx AS ("
    StmtParts(15) = "    SELECT ib.id, ib.idx_name, ib.tbl_name, ii.col0_name, ii.columns"
    StmtParts(16) = "    FROM ib, ii"
    StmtParts(17) = "    ON ib.idx_name = ii.idx_name"
    StmtParts(18) = "),"
    StmtParts(19) = "iex AS ("
    StmtParts(20) = "    SELECT idx.*, il.idx_seq, il.[unique], il.origin, il.partial"
    StmtParts(21) = "    FROM idx, il"
    StmtParts(22) = "    WHERE idx.idx_name = il.idx_name"
    StmtParts(23) = "),"
    StmtParts(24) = "fdup AS ("
    StmtParts(25) = "    SELECT tbl_name, col0_name, count(*) AS group_size"
    StmtParts(26) = "    FROM iex"
    StmtParts(27) = "    WHERE partial = 0"
    StmtParts(28) = "    GROUP BY tbl_name, col0_name"
    StmtParts(29) = "    HAVING group_size > 1"
    StmtParts(30) = "),"
    StmtParts(31) = "idup AS ("
    StmtParts(32) = "    SELECT iex.*, fdup.group_size"
    StmtParts(33) = "    FROM iex"
    StmtParts(34) = "    JOIN fdup"
    StmtParts(35) = "    ON iex.tbl_name = fdup.tbl_name AND iex.col0_name = fdup.col0_name"
    StmtParts(36) = ")"
    StmtParts(37) = "SELECT *"
    StmtParts(38) = "FROM idup AS similar_indices"
    StmtParts(39) = "ORDER BY tbl_name, col0_name, columns"
    
    SimilarIndices = Join(StmtParts, vbNewLine)
End Function
