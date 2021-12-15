---
layout: default
title: Overview
nav_order: 1
parent: SQLiteC
permalink: /sqlitec/overview
---

The structure and functionality of the SQLiteC subpackage shown in [Fig. 1](#SQLiteCCore) are similar to that of [SQLiteADO-Core][]. Since SQLiteC uses the C-API directly, the SQLiteC must implement the functionality provided by SQLiteODBC/ADODB to the LiteADO class. SQLiteCConnection, SQLiteCStatement, SQLIteCParameters, and SQLIteCExecSQL are the four core classes of the SQLiteC package. Additionally, there are three supporting classes, SQLiteC, SQLiteCRC, SQLiteCRecordSetADO.

<a name="SQLiteCCore"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteCCore.svg" alt="SQLiteCCore" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteC vs. SQLiteADOCore</b></p>  

Individual methods provided by these classes can be classified as either SQLite API wrappers, using the SQLite APIs directly, or convenience methods, relying on the API wrappers to provide higher-level functionality. Further, because the abstraction level differs between C and VBA, it often makes sense to have a single VBA method, wrapping a group of related SQLite interfaces.

### The Core

The  SQLite library operates several different [objects][SQLite Objects]. The two most important objects are the [Connection/sqlite3][dbc] and [Statement/sqlite3_stmt][dbs]. The SQLiteCConnection class wraps the Connection object and provides related functionality. The SQLiteCStatement class wraps the Statement object together with its constructors and destructor. Three other core classes - SQLiteCParameters, SQLiteCExecSQL, and SQLiteCMeta - use the Statement object and wrap APIs for handling parameterized queries, executing statements, and retrieving the data.

### Database Manager

The SQLiteC class primarily acts as a manager object and the main entry point for an application. It coordinates setup and teardown functionality and reduces coupling between other components while abstracting away some boilerplate code. SQLiteC also provides a method wrapping the [Backup/sqlite3_backup][backup] object and associated interfaces. This method, however, may need to be moved from SQLiteC, as it takes SQLiteCConnection objects as its dependencies.

### SQLiteCRC

The SQLiteCRC class is a simple dictionary-based container. It keeps the SQLite error codes enum and provides code to code name mapping.

### SQLiteCRecordsetADO

SQLiteCRecordsetADO fabricates an ADODB.Recordset object, which is necessary to implement the ILiteADO interface. SQLiteCExecSQL object instantiates SQLiteCRecordsetADO, injecting SQLiteCStatement into its factory as a dependency. The factory uses the DbExecutor member to retrieve rowset associated metadata from SQLite, translates it to ADODB.Field attributes, assembles Recordset structure, and returns it to the calling SQLiteCExecSQL object for data injection.


<!-- References -->

[SQLiteADO-Core]: ../sqliteado-core#SQLiteADO-Core
[SQLite Objects]: https://www.sqlite.org/c3ref/objlist.html
[dbc]: https://www.sqlite.org/c3ref/sqlite3.html
[dbs]: https://www.sqlite.org/c3ref/stmt.html
[backup]: https://www.sqlite.org/backup.html
