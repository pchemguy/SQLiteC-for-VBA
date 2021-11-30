---
layout: default
title: Manager
nav_order: 1
parent: SQLiteC
permalink: /sqlitec/sqlitec
---

SQLiteC is the top API class of the SQLiteC subpackage. Since this subpackage uses the C-API directly, its members must implement the functionality provided by the SQLiteODBC driver and the ADODB library. While the position of the SQLiteC class is similar to that of LiteMan, the two managers have significant differences.

The SQLiteC class

* generates its instances (factory function),
* calls the [DllManager][] class to load SQLite3 DLL library,
* calls LiteFSCheck to run initial diagnostics and resolve database pathname,
* generates SQLiteCConnection instances (abstract factory),
* initiates resolution of [circular references][] at the termination stage.

SQLiteC also provides a wrapper for the SQLite backup APIs. Since the SQLite version API does not require a database connection, SQLiteC wraps it, providing a basic SQLite3.dll usability check. The *immediate pane* command `?SQLiteC(vbNullString).Version` should print the library version as an integer number, e.g., 3037000. Alternatively, this command `?SQLiteC(vbNullString).Version(False)` prints the same value as a string, e.g., 3.37.0.

Project repository includes custom-compiled SQLite3 binaries for Windows. The Library/SQLiteCAdo/dll folder contains x32 (built with Microsoft Build Tools) and x64 (built with MSYS/MinGW) versions. Both have all optional extensions enabled, including the ICU extension, which depends on the ICU binaries (also included). Further information regarding the building process, including the scripts, is available from the  [SQLite ICU MinGW][] project.

<a name="SQLiteCCore"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteCCore.svg" alt="SQLiteCCore" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteC vs. SQLiteADOCore</b></p>  

The structure and functionality of the SQLiteC subpackage shown in [Fig. 1](#SQLiteCCore) are similar to that of [SQLiteADO-Core][]. The bulk of the classes implement the LiteADO functionality provided by SQLiteODBC/ADODB. At the same time, SQLiteC implements some functional logic differently.

For instance, the SQLiteC constructor does not instantiate a connection object as opposed to the LiteMan class. The latter generates a LiteMan instance with an encapsulated ILiteADO object associated with an already opened ADODB.Connection. SQLiteC, on the other hand, acts essentially as an abstract factory for the SQLiteCConnection class via its SQLiteC.CreateConnection factory. This method has the same signature as the LiteMan factory and a similar logic of calling the LiteFSCheck class, which resolves provided database path before SQLiteC passes it to the SQLiteCConnection factory.

Another difference involves the number of supported connection objects per manager. LiteMan encapsulates a single instance of LiteADO/ILiteADO, which, in turn, encapsulates a single ADODB.Connection. SQLiteC, on the other hand, wraps a Dictionary object used for managing a collection of connection objects. However, since the database path acts as the key, 'main' databases must be unique among connections. One anonymous (such as an in-memory) database (keyed with an empty string) is also acceptable.


<!-- References -->

[DllManager]: https://pchemguy.github.io/DllTools/
[Circular References]: https://pchemguy.github.io/ObjectStore/
[SQLite ICU MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
[SQLiteADO-Core]: ../sqliteado-core#SQLiteADO-Core
