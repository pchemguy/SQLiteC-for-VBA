---
layout: default
title: Manager
nav_order: 1
parent: SQLiteC
permalink: /sqlitec/sqlitec
---

SQLiteC is the top API class of the SQLiteC subpackage. Since this subpackage uses the C-API directly, its members must implement the functionality provided by the SQLiteODBC driver and the ADODB library. While the position of the SQLiteC class is similar to that of LiteMan, the two managers have significant differences.

Project repository includes custom-compiled SQLite3 binaries for Windows. The Library/SQLiteCAdo/dll folder contains x32 (built with Microsoft Build Tools) and x64 (built with MSYS/MinGW) versions. Both have all optional extensions enabled, including the ICU extension, which depends on the ICU binaries (also included). Further information regarding the building process, including the scripts, is available from the  [SQLite ICU MinGW][] project.

The SQLiteC class is responsible for loading the SQLite3.dll library, creating/managing connection objects, and providing a wrapper for the SQLite backup APIs. Additionally, the SQLite version API does not require a database connection, so SQLiteC also wraps this API providing a simple way to check that the SQLite3.dll is usable. The *immediate pane* command `?SQLiteC(vbNullString).Version` should print the library version as an integer number, e.g., 3037000. Alternatively, this command `?SQLiteC(vbNullString).Version(False)` prints the same value as a string, e.g., 3.37.0.

The rest of SQLiteC classes collectively implement functionality

<a name="SQLiteCCore"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteCCore.svg" alt="SQLiteCCore" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteC vs. SQLiteADOCore</b></p>  

[SQLiteADO-Core][] 

The SQLiteC constructor does not instantiate a connection object, as opposed to the LiteMan class. The latter generates a LiteMan instance with an encapsulated ILiteADO object associated with an already opened ADODB.Connection object. SQLiteC, on the other hand, acts essentialy as an abstract factory for SQLiteCConnection class via its SQLiteC.CreateConnection factory. This method has the same signature as the LiteMan factory, and a similar logic of calling the LiteFSCheck to resolve provided database path before passing it to the SQLiteCConnection factory.

While the returned LiteMan instance is already associated with an open Connection object (via LiteADO), the constructor of SQLiteC does not create a connection object. Presently, LiteMan instantiates connections via explicit calls only. Another difference involves the number of supported connection objects per manager. LiteMan encapsulates a single instance of LiteADO, which, in turn, encapsulates a single ADODB.Connection object. SQLiteC, on the other hand, wraps a Dictionary object used for managing a collection of connection objects. However, since the database path acts as the key, the same database file cannot be the 'main' database on more than one connection. One anonymous database keyed with an empty string is also acceptable.


<!-- References -->

[SQLite ICU MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
[SQLiteADO-Core]: ../sqliteado-core#SQLiteADO-Core
