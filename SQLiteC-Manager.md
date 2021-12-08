---
layout: default
title: Manager
nav_order: 2
parent: SQLiteC
permalink: /sqlitec/sqlitec
---

SQLiteC is the top API class of the SQLiteC subpackage. While the position of the SQLiteC class is similar to that of LiteMan, the two managers have some differences.

The SQLiteC class

  1) generates its instances (factory function),
  2) calls the [DllManager][] class to load SQLite3 DLL library,
  3) calls [LiteFSCheck][] to run initial diagnostics and resolve database pathname,
  4) generates SQLiteCConnection instances (abstract factory),
  5) initiates resolution of [circular references][] at the termination stage.

While functions (1), (3), and (4) are similar to those provided by LiteMan, (2) and (5) are specific to SQLiteC. It also provides a wrapper for the SQLite backup APIs. Since the SQLite version API does not require a database connection, SQLiteC wraps it, providing a basic SQLite3.dll usability check. The *immediate pane* command `?SQLiteC(vbNullString).Version` should print the library version as an integer number, e.g., 3037000. Alternatively, this command `?SQLiteC(vbNullString).Version(False)` prints the same value as a string, e.g., 3.37.0.

Project repository includes custom-compiled SQLite3 binaries for Windows. The Library/SQLiteCAdo/dll folder contains x32 (built with Microsoft Build Tools) and x64 (built with MSYS/MinGW) versions. Both have all optional extensions enabled, including the ICU extension, which depends on the ICU binaries (also included). Further information regarding the building process, including the scripts, is available from the  [SQLite ICU MinGW][] project.

When comparing similar functions, there are also certain implementation differences between SQLiteC and LiteMan. For instance, the SQLiteC constructor does not instantiate a connection object as opposed to the LiteMan class. The latter generates a LiteMan instance with an encapsulated ILiteADO object associated with an already opened ADODB.Connection. SQLiteC, on the other hand, acts essentially as an abstract factory for the SQLiteCConnection class via its SQLiteC.CreateConnection factory. This method has the same signature as the LiteMan factory and a similar logic of calling the LiteFSCheck class to resolve provided database path before SQLiteC passes it to the SQLiteCConnection factory.

Another difference involves the number of supported connection objects per manager. LiteMan encapsulates a single instance of LiteADO/ILiteADO, which, in turn, encapsulates a single ADODB.Connection. SQLiteC, on the other hand, wraps a Dictionary object used for managing a collection of connection objects. However, since the database path acts as the key, 'main' databases must be unique among connections. One anonymous (such as an in-memory) database (keyed with an empty string) is also acceptable.

The application using SQLiteAdo should retain its reference for the database manager class LiteMan instance until it discards all library objects. In the case of the SQLiteC subpackage, the application must keep such a reference because termination of the SQLiteC manager initiates the CleanUp cascade turning all manager's descendant objects into unusable garbage.


<!-- References -->

[DllManager]: https://pchemguy.github.io/DllTools/
[LiteFSCheck]: ../integrity
[Circular References]: https://pchemguy.github.io/ObjectStore/
[SQLite ICU MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
