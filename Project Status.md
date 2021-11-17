---
layout: default
title: Performance and known issues
nav_order: 4
permalink: /project-status
---

### Performance considerations: SQLiteADO vs. SQLiteC and know issues

#### Locked database state handling

The first issue is associated with the ADODB library and the SQLiteODBC driver,  and thus, it affects the SQLiteADO subpackage. I discovered this problem while integrating the [SecureADODB fork][] into my demo app [ContactEditor][], and it was the reason I started this project, hoping to find a workaround. This issue occurs when an ADODB object attempts to execute a modifying query (e.g., journal mode change pragma) against a locked database. As a result, the host application (in my case Excel) hangs for 100 s before raising the 'Database is busy' error (see demo macro DemoHostFreezeWithBusyDb in module LiteExamples located in SQLite/ADO/ADemo). The issue may manifest itself due to several different circumstances. The simplest way to reproduce it is to lock the database by starting a transaction using a GUI tool, such as [DB Browser for SQLite][]. LiteFSCheck class attempts to detect various potential issues before opening the database or an attempt to modify it. LiteFSCheck checks for the presence of a pending transaction indirectly via the journal files. SQLiteC subpackage uses the SQLiteC API directly and is not affected by this issue. SQLite returns the 'Database is busy' status via the API right away, and SQLiteC can handle it without hanging the application.

#### Circular references

The other issue affects the SQLiteC subpackage. When I was satisfied with the initial draft of the library, I switched to preparing documentation, including the diagrams. I recognized the circular reference pattern on the class diagram ([Fig. 2][SQLiteC classes]) and recalled reading the [Lazy Object / Weak Reference][Weak Reference] post on the RDVBA blog. My first approach to disentangle this circular reference web was the introduction of an explicit cleanup cascade. Since SQLiteC is at the top of the hierarchy and is not involved in circular references, VBA should handle its destruction correctly, and its Class_Terminate should be a suitable place to start this cascade. This cascade descends to the lowest affected level and sets parent references to nothing as it ascends back to the top. With this cascade enabled, the following issue occurs. Attempting to exit the host application (in my case Excel) causes the application to start using CPU actively and hang (often the usage level jumps to and remains at 100% of the CPU core / HT until the process is killed). This issue can be triggered by running the entire test suite. I identified several tests in SQLiteCTests as possible culprits and commented them out. The remaining test suite is still affected (possibly after the addition of new tests). Curiously, if I run the suspect tests individually, the issue does not occur but only manifests itself when certain test sets, including the suspects, are run. The problem may still be fixable by tuning the cleanup cascade. Or this approach may be problematic to start with, as mentioned in the comments to the post referenced above. The alternative approach relies on weak references, as discussed in the [post][Weak Reference], but it likely has its cons as well.

#### Relative performance

The third issue is related to the performance of DLL calls from VBA. It should affect both subpackages, as SQLiteC calls the SQLite DLL explicitly, and SQLiteADO calls the SQLiteODBC driver via the ADODB library. I will provide further details on a separate page. Initial rough performance tests suggest that SQLiteC is more efficient than pure ADODB with scalar queries, but ADODB outperforms SQLiteC when a set of rows is retrieved. This result is very preliminary, however. While coding the SQLiteC package, I kept in mind efficiency consideration overall but have not made any profiling tuning yet, postponing this process until a properly working draft is available. I should also add the [SQLiteForExcel][] project, which wraps SQLite C API in a regular module, as a performance reference.


<!-- References -->

[SecureADODB fork]: https://pchemguy.github.io/SecureADODB-Fork/
[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[DB Browser for SQLite]: https://sqlitebrowser.org/
[SQLiteC classes]: /SQLite-C-API/class-hierarchy#SQLiteC
[Weak Reference]: https://rubberduckvba.wordpress.com/2018/09/11/lazy-object-weak-reference/
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
