---
layout: default
title: LiteMan - database manager
parent: SQLiteADO core
nav_order: 2
permalink: /sqliteado/liteman
---

The *LiteMan* class is the top-level API class. This predeclared class generates its instances via the Create factory. The factory is the default class member taking a string describing the database (relative/full path or a special name) and an optional argument controlling whether a new database is acceptable and whether *LiteFSCheck* should use its path resolution protocol. If no error occurs, *LiteMan* passes the resolved database descriptor to the *LiteADO* class, which generates the *ILiteADO* database object. *LiteMan* also provides routines for checking driver availability, attaching/detaching/cloning databases, placing the query result on an Excel worksheet, and controlling the journal type.

The following table shows some commands, which can be executed from the *immediate pane* without any additional code:

<p align="center"><b>Table 1. Sample immediate pane commands</b></p>

| |                 Command                   |                Output                 | |
|-|-------------------------------------------|---------------------------------------|-|
| | `?LiteMan.SQLite3ODBCDriverCheck()`       | *True*, if the ODBC driver is found.  | |
| | `?LiteMan(":mem:").ExecADO.GetScalar("")` | SQLite version number, e.g. *3.32.3*  | |
| | `?LiteMan(":tmp:").ExecADO.MainDB`        | Path to the new db in the Temp folder | |
| |                                           |                                       | |
