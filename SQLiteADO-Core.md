---
layout: default
title: SQLiteADO core
nav_order: 5
permalink: /sqliteado-core
---

### Database connectivity

The main classes in this group are LiteADO and LiteMan. ILiteADO has a place of its own: it is an interface class, which does not implement any functionality, but it enables the modular design of the library. Together with the supporting ADOlib class and a module with usage examples, these classes are located in the SQLite/ADO folder.

<a name="SQLiteADO-Core"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO-Core.svg" alt="SQLiteADO Core" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO core classes diagram</b></p>  

### LiteMan

From the calling code's perspective, the LiteMan class is the top-level API object. Shown in the blue rectangle, this database manager class serves to simplify common workflows, and so it is a convenience tool rather than an essential component. LiteMan is a predeclared class generating its instances via the Create factory. The factory is the default class member taking a string describing the database (relative/full path or a special name) and an optional argument controlling whether a new database is acceptable and whether LiteFSCheck should use its path resolution protocol. If no error occurs, LiteMan passes the resolved database descriptor to the LiteADO class, which generates the ILiteADO database object. LiteMan also provides routines for checking driver availability, attaching/detaching/cloning databases, placing the query result on an Excel worksheet, and controlling the journal type. The following table shows some commands, which can be executed from the *immediate pane* without any additional code:

<p align="center"><b>Table 1. Sample immediate pane commands</b></p>

|                 Command                   |                Output                  |  
|-------------------------------------------|----------------------------------------|  
| `?LiteMan.SQLite3ODBCDriverCheck()`       | **True**, if the ODBC driver is found. |  
| `?LiteMan(":mem:").ExecADO.GetScalar("")` | SQLite version number, e.g. **3.32.3** |  
| `?LiteMan(":tmp:").ExecADO.MainDB`        | Path to the new db in the Temp folder  |  
|                                           |                                        |  

### LiteADO

The top-level class of the package core is the LiteADO class, which is the only class that interacts with the ADODB library directly (note the green rectangles with ADODB objects in [Fig. 1](#SQLiteADO-Core) associated with LiteADO). It generates ILiteADO/LiteADO database objects and implements high-level methods for interacting with SQLite databases via the ADODB library and SQLiteODBC driver. The LiteADO's constructor instantiates a new ADODB.Command object and sets its ActiveConnection property to the connection string, thus, opening the database connection. Typically, it remains open for the lifetime of the ILiteADO/LiteADO object. The constructor also sets the default SQL query to the statement querying the SQLite version.

As mentioned earlier, this class is a feature-limited ADODB wrapper. Presently, it does not process events from the ADODB objects, handle parameterized queries (as opposed to SQLiteC), or handle ADODB errors.

### ILiteADO

ILiteADO formalizes LiteADO's interface, and it represents the main database-associated object. ILiteADO provides a unified high-level interface for interactions with SQLite databases. It provides an interface for opening/closing its database connection and methods for executing queries and controlling transactions. The two other groups of SQLiteADO classes take an ILiteADO object as a dependency. Originally a part of the SQLiteADO package, those classes can now be used just as fine by SQLiteC via its own implementation of ILiteADO.

### ADOlib

ADOlib is a legacy module designed as a container for helper routines involving operations on ADODB objects. Presently, SQLiteCAdo uses only one method, RecordsetToQT, which outputs the contents of a Recordset onto an Excel worksheet via the QueryTable feature (see SQLiteIntropectionExample in 'SQLite/MetaSQL/Examples' for examples).
