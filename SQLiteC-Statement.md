---
layout: default
title: Statement
nav_order: 4
parent: SQLiteC
permalink: /sqlitec/statement
---

#### SQLite API wrappers

* Statement meta  
  *Busy()* and *ReadOnly()* provide information on whether a prepared statement 1) has been reset or 2) modifies the database.  
  *SQLQueryOriginal()* and *SQLQueryExpanded* provide query information.
  *PreparedStatementsGet()* enumerates non-finalized prepared statements and either returns them in a dictionary with their statuses or finalizes them. This method should be called on the default class instance.  
* Statement lifecycle control  
  *Prepare16V2()*, *Reset()*, and *Finalize()* - constructor, resetter, and destructor.  

#### User-controlled header fields

*ApplicationId* and *UserVersion* accessors retrieve and set two user-controlled header fields via high-level wrappers.

#### Query execution

* *ExecuteSetup()* is a helper routine. The rest of the query routines call *ExecuteSetup()* for statement preparation and parameter binding. They do not finalize the statement object, and *ExecuteSetup()* will reuse a parameterized prepared statement if a blank query is provided on subsequent calls.
* *ExecuteNonQuery()* executes an SQL statement, possibly parameterized, not returning data. For non-parameterized (plain) queries, SQLiteCConnection.ExecuteNonQueryPlain() interface is used.
* *GetScalar()* executes an SQL statement, possibly parameterized, returning a scalar value. The actual query may return more than one row/column, but only the first field in the first row is returned, discarding the rest.  
* *GetPagedRowSet()* executes an SQL statement, possibly parameterized, returning a set of rows. The returned result is a 1D array of pages, with each non-empty element being a 1D array of rows and each row being a 1D array of field values.
* *GetRowSet2D()* executes an SQL statement, possibly parameterized, returning a set of rows. The result is a row-wise 2D array.
* *GetRecordset()* executes an SQL statement, possibly parameterized, returning a set of rows. The result is a fabricated ADODB.Recordset object.

#### ILiteADO interface

SQLiteCStatement class implements ILiteADO interface. While the current LiteADO/ILiteADO implementation does not support parameterized queries, the SQLiteCStatement/ILiteADO implementation provides such support. Another important aspect is that this implementation handles the SQLite connection object. If a connection is not opened before one of the query methods is called, it is opened and closed; otherwise, it remains open.
