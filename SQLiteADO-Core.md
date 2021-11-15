---
layout: default
title: SQLiteADO core
nav_order: 4
permalink: /sqliteado-core
---

### Database connectivity: *LiteADO* and *LiteMan*

From the calling code's perspective, the LiteMan class is the top-level API object. Shown in the blue rectangle, this database manager class serves to simplify common workflows, and so it is a convenience tool rather than an essential component.

The top-level class of the package core is the LiteADO class, which is the only class that interacts with the ADODB library directly (note the green rectangles with ADODB objects in [Fig. 1](#SQLiteADO-Core) associated with LiteADO). ILiteADO formalizes LiteADO's interface, and it represents the main database-associated object. All other core classes encapsulate the ILiteADO object, if necessary. Therefore, the remaining components can be used by the SQLiteC package as well via its ILiteADO implementation.

<a name="SQLiteADO-Core"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO-Core.svg" alt="SQLiteADO Core" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO core classes diagram</b></p>  

### LiteMan

The LiteMan class is the top-level API class. This predeclared class generates its instances via the Create factory. The factory is the default class member taking a string describing the database (relative/full path or a special name) and an optional argument controlling whether a new database is acceptable and whether LiteFSCheck should use its path resolution protocol. If no error occurs, LiteMan passes the resolved database descriptor to the LiteADO class, which generates the ILiteADO database object. LiteMan also provides routines for checking driver availability, attaching/detaching/cloning databases, placing the query result on an Excel worksheet, and controlling the journal type. The following table shows some commands, which can be executed from the *immediate pane* without any additional code:

<p align="center"><b>Table 1. Sample immediate pane commands</b></p>

  |                 Command                   |                Output                  |  
  |-------------------------------------------|----------------------------------------|  
  | `?LiteMan.SQLite3ODBCDriverCheck()`       | **True**, if the ODBC driver is found. |  
  | `?LiteMan(":mem:").ExecADO.GetScalar("")` | SQLite version number, e.g. **3.32.3** |  
  | `?LiteMan(":tmp:").ExecADO.MainDB`        | Path to the new db in the Temp folder  |  
  |                                           |                                        |  

