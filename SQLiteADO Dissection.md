---
layout: default
title: SQLiteADO dissection
nav_order: 4
permalink: /sqliteado-dissection
---

### SQLiteADO functional structure

SQLiteADO includes a basic wrapper around the ADODB library. SQLiteADO classes can be grouped into several sets based on their functionality.

#### Database connectivity: *LiteADO* and *LiteMan*

From the calling code's perspective, the LiteMan class is the top-level API object. Shown in the blue rectangle, this database manager class serves to simplify common workflows, and so it is a convenience tool rather than an essential component.

The top-level class of the package core is the *LiteADO* class, which is the only class that interacts with the ADODB library directly (note the green rectangles with ADODB objects in [Fig. 1](#SQLiteADO) associated with LiteADO). *ILiteADO* formalizes LiteADO's interface, and it represents the main database-associated object. All other core classes encapsulate the ILiteADO object, if necessary. Therefore, the remaining components can be used by the SQLiteC package as well via its ILiteADO implementation.

#### Validation and integrity: *LiteACID*, *LiteFSCheck*, and *CommonRoutines*

*LiteACID* and *LiteFSCheck* classes in the middle of the figure with the CommonRoutines module validate connection settings and consistency/integrity related functionality.

#### Metadata: *LiteMetaADO*, *LiteMetaSQL*, and *LiteMetaSQLIdxFK*

The three *LiteMeta-* classes at the bottom are focused on providing SQLite introspection functionality. *LiteMetaSQL* and *LiteMetaSQLIdxFK* generate SQL statements used to retrieve metadata about both the engine and the database. *LiteMetaSQL* generates basic queries, and *LiteMetaSQLIdxFK* generates specially crafted bulky SQL queries yielding extended information about foreign keys and indices, as discussed later. *LiteMetaADO* executes select *LiteMetaSQL* queries using an ILiteADO instance and returns actual metadata, facilitating an SQL-based database clone process provided by a member of the *LiteMan* class.

<a name="SQLiteADO"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO.svg" alt="SQLiteADO" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO class diagram</b></p>  
