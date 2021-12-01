---
layout: default
title: Performance and known issues
nav_order: 4
permalink: /project-status
---

### Performance considerations: SQLiteADO vs. SQLiteC and know issues

A few general aspects of the SQLiteC subpackage design could be improved. The SQLiteCConnection class currently incorporates several groups of functions. While I used the ADODB.Connection class as a reference, SQLiteCConnection appears overloaded and should benefit from refactoring and splitting. The functionality incorporated into the SQLiteC class is more focused. However, the backup routine should probably be moved to SQLiteCConnection. The other manager, LiteMan, clearly needs to be refactored. While I made a few attempts to reduce the coupling of SQLiteCAdo classes, and SQLiteADO subpackage is largely decoupled, reducing the coupling of SQLiteC classes could be beneficial.

The design of the SQLiteC package also incorporates several circular reference loops ([Fig. 2][SQLiteC classes]). I only realized this matter once I drafted the class diagram. This topic is discussed in more detail [here][ObjectStore], and the current implementation of the SQLiteC package resolves circular references via a CleanUp cascade.

Although preliminary tests suggest the current SQLiteCAdo performance is reasonable, I have not evaluated it carefully nor optimized it. While there is room for improvement, it is too early to invest efforts into profiling before refactoring discussed above is performed.

#### Locked database state handling

The first issue is associated with the ADODB library and the SQLiteODBC driver,  and thus, it affects the SQLiteADO subpackage. I discovered this problem while integrating the [SecureADODB fork][] into my demo app [ContactEditor][], and it was the reason I started this project, hoping to find a workaround. This issue occurs when an ADODB object attempts to execute a modifying query (e.g., journal mode change pragma) against a locked database. As a result, the host application (in my case Excel) hangs for 100 s before raising the 'Database is busy' error (see demo macro DemoHostFreezeWithBusyDb in module LiteExamples located in SQLite/ADO/ADemo). The issue may manifest itself due to several different circumstances. The simplest way to reproduce it is to lock the database by starting a transaction using a GUI tool, such as [DB Browser for SQLite][]. LiteFSCheck class attempts to detect various potential issues before opening the database or an attempt to modify it. LiteFSCheck checks for the presence of a pending transaction indirectly via the journal files. SQLiteC subpackage uses the SQLiteC API directly and is not affected by this issue. SQLite returns the 'Database is busy' status via the API right away, and SQLiteC can handle it without hanging the application.

#### Relative performance

The third issue is related to the performance of DLL calls from VBA. It should affect both subpackages, as SQLiteC calls the SQLite DLL explicitly, and SQLiteADO calls the SQLiteODBC driver via the ADODB library. I will provide further details on a separate page. Initial rough performance tests suggest that SQLiteC is more efficient than pure ADODB with scalar queries, but ADODB outperforms SQLiteC when a set of rows is retrieved. This result is very preliminary, however. While coding the SQLiteC package, I kept in mind efficiency consideration overall but have not made any profiling tuning yet, postponing this process until a properly working draft is available. I should also add the [SQLiteForExcel][] project, which wraps SQLite C API in a regular module, as a performance reference.


<!-- References -->

[SQLiteC classes]: /SQLite-C-API/class-hierarchy#SQLiteC
[ObjectStore]: https://pchemguy.github.io/ObjectStore/
[SecureADODB fork]: https://pchemguy.github.io/SecureADODB-Fork/
[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[DB Browser for SQLite]: https://sqlitebrowser.org/
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
