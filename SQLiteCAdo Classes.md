---
layout: default
title: Class hierarchy
nav_order: 3
permalink: /class-hierarchy
---

### SQLiteADO subpackage

I started the SQLiteADO package shown in [Fig. 1](#SQLiteADO) as a small VBA-SQLite project, which included

* introspection features,
* SQLIteODBC connection string related routines, and
* database integrity/clone functionality.

<a name="SQLiteADO"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteADO.svg" alt="SQLiteADO" width="100%" /></div>
<p align="center"><b>Fig. 1. SQLiteADO class diagram</b></p>  

I planned to integrate this package with the [SecureADODB][] library and implemented the core functionality as a single VBA class with a few supporting SQL generating classes. To avoid cyclic dependency with SecureADODB, I included only very basic ADODB wrappers, which is why the present version of this package is feature-limited. Currently, SQLiteADO does not
 
* process events from the ADODB objects,
* handle parameterized queries,
* handle ADODB errors.

I still consider integration with SecureADODB in the future, so for now, this subpackage will remain a working prototype. In its present state, its classes can be grouped into several sets based on their functionality, including *database connectivity (the core), validation/integrity, and metadata*.

### SQLiteC subpackage

When I realized that some limitations of the SQLiteODBC driver are difficult, if not impossible, to overcome, I started looking into alternative SQLite access options. While the SQLiteODBC driver is open source, I did not want to deal with it at that point, so I focused on the [SQLiteForExcel][] project. Its demo code worked out of the box with both VBA6/x32 and VBA7/x64 environments, and it should be possible to use a recent official x32 SQLite binary via the included adaptor. However, I wanted to use custom compiled SQLite binaries anyway to enable omitted features. This way, I could remove the additional indirection layer due to the adaptor and make the two versions of the code (x32 and x64) more similar. Another reason for a new project was the lack of the higher-level API, which made it necessary to work with the low-level wrappers directly. However, having the entire codebase in a single code module did not facilitate its use. For these reasons, I started developing an object-oriented SQLiteC package ([Fig. 2](#SQLiteC)). While I coded SQLiteC from scratch, I borrowed some API declarations from *SQLiteForExcel* and used its supporting code as a reference.

<a name="SQLiteC"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/SQLiteC.svg" alt="SQLiteC" width="100%" /></div>
<p align="center"><b>Fig. 2. SQLiteC class diagram</b></p>  

While developing the two packages, SQLiteADO (formerly SQLiteDB) and SQLiteC, I realized that it would be instructive to combine the two packages and define an interface class unifying and formalizing their high-level APIs. Thus, the two packages formed the *SQLite C/ADO with Introspection for VBA* library.


<!-- References -->

[SecureADODB]: https://rubberduckvba.wordpress.com/2020/04/22/secure-adodb/
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
