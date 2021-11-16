---
layout: default
title: Class hierarchy
nav_order: 3
permalink: /class-hierarchy
---

### SQLiteADO subpackage

I conceived this SQLite-VBA project as [SQLiteDB][] while integrating the [SecureADODB][] library into my demo Excel/VBA-based application [ContactEditor][]. I chose SQLite as the primary backend and [DB Browser for SQLite][] as my primary GUI administration tool. Soon I realized, however, that due to the nature of SQLite, Excel used the copy of SQLite embedded into the SQLiteODBC driver, limiting the usefulness of DB Browser for SQLite. At the same time, due to the modular architecture of SQLite, many features are only available if enabled during compilation. Therefore, it is prudent to probe available functionality, and executing queries via the same ADODB-SQLiteODBC path appeared the most straightforward way of obtaining relevant information. For this reason, I started developing a set of routines generating introspection SQL queries. Another feature I wanted to incorporate into this library was an ADO connection string helper; later, I added database cloning and integrity-related functionality.

I planned to integrate SQLiteDB with my fork of [SecureADODB][SecureADODB PG] and use it as an intermediary between the application and SecureADODB. On the other hand, online integrity features, for example, required an active database connection. To avoid cyclic dependency between the two libraries, I added to SQLiteDB a few feature-limited ADODB wrappers to satisfy its needs. Eventually, I realized that it was time for refactoring the code base, yielding the SQLiteADO package illustrated in [Fig. 1](#SQLiteADO).

<a name="SQLiteADO"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO.svg" alt="SQLiteADO" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO class diagram</b></p>  

While I significantly refactored the code and added some features, the package is still primarily focused on

* database connectivity (ADO/SQLiteODBC connection string helpers and limited ADODB wrappers),
* validation/integrity, and
* metadata (SQL-based SQLite introspection).

Currently, the ADODB wrapper module of SQLiteADO does not handle

* events from the ADODB objects,
* parameterized queries,
* ADODB errors.

### SQLiteC subpackage

When I realized that some limitations of the SQLiteODBC driver are difficult, if not impossible, to overcome, I started looking into alternative SQLite access options. While the SQLiteODBC driver is open source, I did not want to deal with it at that point, so I focused on the [SQLiteForExcel][] project. Its demo code worked out of the box with both VBA6/x32 and VBA7/x64 environments, and it should be possible to use a recent official x32 SQLite binary via the included adaptor. However, I wanted to use custom compiled SQLite binaries anyway to enable omitted features. This way, I could remove the additional indirection layer due to the adaptor and make the two versions of the code (x32 and x64) more similar. Another reason for a new project was the lack of the higher-level API, which made it necessary to work with the low-level wrappers directly. However, having the entire codebase in a single code module did not facilitate its use. For these reasons, I started developing an object-oriented SQLiteC package ([Fig. 2](#SQLiteC)). While I coded SQLiteC from scratch, I borrowed some API declarations from *SQLiteForExcel* and used its supporting code as a reference.

<a name="SQLiteC"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteC.svg" alt="SQLiteC" width="100%" /></div>
<p align="center"><b>Fig. 2. SQLiteC class diagram</b></p>  

While developing the two packages, SQLiteADO and SQLiteC, I realized that it would be instructive to combine the two packages and define an interface class unifying and formalizing their high-level APIs. Thus, the two packages formed the *SQLite C/ADO with Introspection for VBA* library.


<!-- References -->

[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[SecureADODB]: https://rubberduckvba.wordpress.com/2020/04/22/secure-adodb/
[DB Browser for SQLite]: https://sqlitebrowser.org/
[SQLiteDB]: https://pchemguy.github.io/SQLiteDB-VBA-Library/
[SecureADODB PG]: https://pchemguy.github.io/SecureADODB-Fork/
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
