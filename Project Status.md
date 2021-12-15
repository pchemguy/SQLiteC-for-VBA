---
layout: default
title: Performance and known issues
nav_order: 6
permalink: /project-status
---

### Performance considerations: SQLiteADO vs. SQLiteC and know issues

A few general aspects of the SQLiteC subpackage design could be improved. The SQLiteCConnection class currently incorporates several groups of functions. While I used the ADODB.Connection class as a reference, SQLiteCConnection appears overloaded and should benefit from refactoring and splitting. The functionality incorporated into the SQLiteC class is more focused. However, the backup routine should probably be moved to SQLiteCConnection. The other manager, LiteMan, clearly needs to be refactored. While I made a few attempts to reduce the coupling of SQLiteCAdo classes, and SQLiteADO subpackage is largely decoupled, reducing the coupling of SQLiteC classes could be beneficial.

I still have occasional issues with DllManager leaking resources. As a result, DllManager cannot be load SQLite DLL, necessitating the restart of Excel.

The design of the SQLiteC package also incorporates several circular reference loops ([Fig. 2][SQLiteC classes]). I only realized this matter once I drafted the class diagram. This topic is discussed in more detail [here][ObjectStore], and the current implementation of the SQLiteC package resolves circular references via a CleanUp cascade.

Although preliminary tests suggest the current SQLiteCAdo performance is reasonable, I have not evaluated it carefully nor optimized it yet. While there is room for improvement, it is too early to invest efforts into profiling before refactoring discussed above is performed.

#### Locked database state handling

Both the ADODB library and SQLiteODBC driver define the timeout feature. The driver, however, has a bug: it ignores the timeout value set by the ADODB objects. The default SQLiteODBC timeout, defined in its source, is 100s, and it can only be changed via the connection string option 'Timeout=XXX;' (XXX - value in ms). Also, if the 'StepAPI=True' option is specified, the timeout feature is disabled completely regardless of the XXX value above. The LiteFSCheck class checks for potential issues before opening the database or executing a modifying query to prevent timeout-related errors. It also checks for the presence of a pending transaction indirectly via the journal files. SQLiteC subpackage uses the SQLiteC API returning the 'Database is busy' status due to a database lock right away, without raising an error or hanging the application.

#### Relative performance

Another significant issue is the [performance of DLL calls][DLL calls] from VBA. The issue does not affect Excel 2002-x32, but most VBA7-based Office environments are probably affected by this Microsoft [AMSI][] penalty. What is now relatively clear from the referenced page is why calling an empty DLL routine (no arguments, no return value) takes 8ns under Excel 2002/VBA6/x32 and 2us 2016/VBA7/x64.

AMSI penalty should affect both subpackages, as SQLiteC calls the SQLite DLL explicitly, and SQLiteADO calls the SQLiteODBC driver via the ADODB library. Initial rough performance tests suggest that SQLiteC is more efficient than pure ADODB with scalar queries, but ADODB may outperform SQLiteC when a set of rows is retrieved. This result is very preliminary, however. While coding the SQLiteC package, I kept in mind efficiency consideration overall but have not made any profiling tuning yet, postponing this process until a properly working draft is available. I should also add the [SQLiteForExcel][] project, which wraps SQLite C API in a regular module, as a performance reference.


<!-- References -->

[SQLiteC classes]: ./class-hierarchy#SQLiteC
[ObjectStore]: https://pchemguy.github.io/ObjectStore/
[SecureADODB fork]: https://pchemguy.github.io/SecureADODB-Fork/
[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[DB Browser for SQLite]: https://sqlitebrowser.org/
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
[DLL calls]: https://pchemguy.github.io/DllTools/vba-dll-call
[AMSI]: https://www.microsoft.com/security/blog/2018/09/12/office-vba-amsi-parting-the-veil-on-malicious-macros/#caption-attachment-97305
