---
layout: default
title: SQLiteADO core
nav_order: 4
permalink: /sqliteado-core
---

### Database connectivity: *LiteADO* and *LiteMan*

From the calling code's perspective, the LiteMan class is the top-level API object. Shown in the blue rectangle, this database manager class serves to simplify common workflows, and so it is a convenience tool rather than an essential component.

The top-level class of the package core is the LiteADO class, which is the only class that interacts with the ADODB library directly (note the green rectangles with ADODB objects in [Fig. 1](#SQLiteADO-Core) associated with LiteADO). ILiteADO formalizes LiteADO's interface, and it represents the main database-associated object. All other core classes encapsulate the ILiteADO object, if necessary. Therefore, the remaining components can be used by the SQLiteC package as well via its ILiteADO implementation.

### LiteMan

The LiteMan class is the top-level API class. This predeclared class generates its instances via the Create factory. The factory is the default class member taking a string describing the database (relative/full path or a special name) and an optional argument controlling whether a new database is acceptable and whether LiteFSCheck should use its path resolution protocol. If no error occurs, LiteMan passes the resolved database descriptor to the LiteADO class, which generates the ILiteADO database object. LiteMan also provides routines for checking driver availability, attaching/detaching/cloning databases, placing the query result on an Excel worksheet, and controlling the journal type. The following table shows some commands, which can be executed from the *immediate pane* without any additional code:

<p align="center"><b>Table 1. Sample immediate pane commands</b></p>

|                 Command                   |                Output                  |  
|-------------------------------------------|----------------------------------------|  
| `?LiteMan.SQLite3ODBCDriverCheck()`       | **True**, if the ODBC driver is found. |  
| `?LiteMan(":mem:").ExecADO.GetScalar("")` | SQLite version number, e.g. **3.32.3** |  
| `?LiteMan(":tmp:").ExecADO.MainDB`        | Path to the new db in the Temp folder  |  
|                                           |                                        |  

### LiteADO

The LiteADO class generates ILiteADO/LiteADO database objects and is responsible for interacting with SQLite databases via the ADODB library and SQLiteODBC driver. The LiteADO's constructor instantiates a new ADODB.Command object and sets its ActiveConnection property to the connection string, thus, opening the database connection. Typically, it remains open for the lifetime of the ILiteADO/LiteADO object. The constructor also sets the default SQL query to the statement querying the SQLite version.

As mentioned earlier, this class is a feature-limited ADODB wrapper. Presently, it does not process events from the ADODB objects, handle parameterized queries (as opposed to SQLiteC), or handle ADODB errors.

### ILiteADO

ILiteADO is an interface class representing a database-bound object and providing a unified interface for high-level interactions with SQLite databases. It provides methods for executing queries and controlling transactions. It also provides an interface to open/close its database connection.

The two other groups of SQLiteADO classes take an ILiteADO object. Even though those classes have been developed as a part of the SQLiteADO package, they can now be used just as fine by SQLiteC via its own implementation of ILiteADO.

### ADOlib

ADOlib is a legacy module designed as a container for helper routines involving operations on ADODB objects. Presently, SQLiteCAdo uses only one method, RecordsetToQT, which outputs the contents of a Recordset onto an Excel worksheet via the QueryTable feature (see SQLiteIntropectionExample in 'SQLite/MetaSQL/Examples' for examples).

<a name="SQLiteADO-Core"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO-Core.svg" alt="SQLiteADO Core" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO core classes diagram</b></p>  
