---
layout: default
title: Meta
nav_order: 6
parent: SQLiteC
permalink: /sqlitec/meta
---

SQLite C API provides a set of metadata interfaces (metadata available via SQL queries is beyond the scope of this section) with information dispersed over several pages ([meta1][], [meta2][], [meta3][], [meta4][], and [meta5][]). Importantly, all these interfaces provide information about a single column in a prepared statement (with one exception noted below). There are three sources of metadata (16-bit UTF versions not shown):

* CREATE TABLE DDL statement (table source):
	* sqlite3_table_column_metadata
	* sqlite3_column_database_name
	* sqlite3_column_table_name
	* sqlite3_column_origin_name
	* sqlite3_column_decltype
* SQL statement (query source):
	* sqlite3_column_count
	* sqlite3_column_name
* particular table value (value source):
	* sqlite3_column_type

Table interfaces do not provide meaningful information on generated columns, possibly returning error codes or null pointers (see the official documentation for further details). The *sqlite3_column_name* interface returns the column alias from the query. When no alias is specified, the returned value is unspecified. The *sqlite3_column_type* interface returns type information about a particular row value, which is only meaningful if the database cursor points at a valid row.

SQLite C API supports retrieval of a single value only, i.e., a value in a particular column from the row pointed to by the database cursor (disregarding legacy interfaces). Furthermore, SQLite provides a set of type-specific column_* interfaces for this purpose. The caller is responsible for selecting a particular interface, with suitable selection reducing the chances of undesirable silent automatic typecasting.

SQLite employs dynamic typing. Each column has the so-called *Affinity* type determined based on the declared column type defined in the CREATE DDL statement. Affinity is the preferred column type, meaning that the column may store values of all SQLite supported types regardless of its affinity type. Type affinity is not available for generated columns, and for columns coming from a table, the *sqlite3_column_decltype* interface returns column DDL type definition as a string. *TypeAffinityFromDeclaredType()* converts this string to type affinity following the algorithm from the official SQLite documentation, and the *AffinityMap* attribute facilitates mapping affinity type to SQLite data type. Even for columns storing only values of a single data type matching the column affinity, it is still necessary to check the data type of individual for possible NULL.

*ColumnMetaAPI()* is the main method of the SQLiteCMeta class, retrieving all available metadata. *TableMetaCollect()*, in turn, wraps *ColumnMetaAPI()* and loops through columns. It also presets the ColumnIndex value in the metadata structure passed ByRef to *ColumnMetaAPI()*. The *Initialized* attribute is a flag used to confirm that the caller has set the ColumnIndex attribute. The *MetaLoaded* attribute of the SQLiteCExecSQL class is also a flag indicating that metadata is available.


<!-- References -->

[step API]: https://www.sqlite.org/c3ref/step.html
[meta1]: https://www.sqlite.org/c3ref/column_database_name.html
[meta2]: https://www.sqlite.org/c3ref/table_column_metadata.html
[meta3]: https://www.sqlite.org/c3ref/column_name.html
[meta4]: https://www.sqlite.org/c3ref/column_blob.html
[meta5]: https://www.sqlite.org/c3ref/column_decltype.html
