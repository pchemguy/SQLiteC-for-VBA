---
layout: default
title: Introduction
nav_order: 3
permalink: /
---

### Motivation and scope

First introduced for automating MS Excel, VBA is still the only native embedded programming language available in MS Office. VBA is an object-oriented scripting language requiring no additional software for development (though [RDVBA][] is a great, if not essential, addition). MS Excel, in turn, provides a convenient means for data management, manipulation, and presentation, requiring little to no programming experience. For these reasons, it is appealing to a broad spectrum of users. Various data-focused tasks, both personal and business-related, take advantage of Excel functionality even when an [RDBMS][] database is more suitable.

For years, I have used Excel as the poor man's database and more recently started thinking about migrating the data to a proper database. I could not find an existing personal information manager (PIM) with all desirable features and started a new project. VBA may not be the most natural choice for a new project. But its tight integration with Excel, my data keeper, was a significant reason for going with VBA. If successful, it might be necessary to migrate the project, for example, to Python. From this point of view, I also consider VBA/Excel as a convenient prototyping platform.

VBA incorporates several [technologies][VBA RDBMS] for accessing databases, providing a means for Excel to interact with database engines. These database-related VBA libraries present various levels of abstraction and, depending on the database, may need extra drivers. The [ADODB][] library is a medium-level library designed for use by applications. However, in the common usage scenarios, the entire repertoire of technical features exposed by the library may not be necessary. Therefore, "middleware" packages may wrap and encapsulate unnecessary details, simplifying the client code and workflows.

While VBA can communicate with many RDBMS via the ADODB library, this project focuses on SQLite. There are several reasons for making this choice. SQLite is notable for its simplified serverless architecture, with the entire engine fitting within a single compact dynamic loading library. An application, such as a lightweight data manager, may embed this library or incorporate it as a standalone DLL. Usually, the application and the SQLite engine reside locally, enabling communication via in-process library calls. Therefore, database operations can be performed offline and, potentially, several orders of magnitude faster due to eliminated network and inter-process delays without caching. For these reasons, SQLite can also act as a local cache layer for network databases. In-memory databases supported by SQLite may further improve the speed of database operations, acting as a cache layer. While SQLite may not meet scalability and multi-user demands, its advantages make it well suited for local information management.

### Existing database-centric VBA "middleware" projects

A few existing VBA projects develop database-related tools worth mentioning. Here are three generic projects:

* [SecureADODB][] is a thin wrapper around the ADODB library. SecureADODB is a generic (database-agnostic) fully object-oriented VBA library focused on database connectivity workflows. The project employs advanced OOP techniques and patterns, demonstrating VBA implementation details, and presents powerful testing strategies.
* [VBA SQL Library][] targets the generation of SQL queries, another common component of the database interaction workflow. I have not studied its code yet or tried to use it, so I cannot comment on its functionality.
* [vbaMyAdmin][] focuses on providing a GUI database administration functionality. I have only recently discovered this application, so no further comments here either.

And the following two projects specifically aim at SQLite access from VBA:

* [winsqlite3.dll-4-VBA][] demonstrates the use of the winsqlite3.dll library, a part of Windows distribution, and can be accessed directly from both VBA x32 and x64. While I managed to run the code from the project repository, it kept crashing Excel.
* [SQLite for Excel][] provides a standard module with wrappers around SQLite C-language API. Further, since VBA-x32 cannot call official SQLite3 binaries directly, the project also provides a compiled x32 DLL adaptor. It has both advantages and disadvantages, as further later.

### SQLite

A common means of accessing SQLite databases from VBA is via the ADODB library and the [ChW SQLiteODBC][] driver. The driver embeds a copy of the SQLite library, so no other software is necessary. However, the driver has a few limitations. As of this writing (Fall 2021), ChW has not updated the driver since June 2020, and the embedded copy of SQLite is an outdated feature-limited build. While ChW also provides an experimental version of SQLiteODBC binaries that should work with a local copy of the SQLite library, I could not make it work. The alternative options include using a [custom-built][SQLiteODBC PG] SQLiteODBC driver or bypassing it entirely. This project explores both of them.

Bypassing SQLiteODBC/ADODB relies on the VBA's DLL calling ability. This feature enables, for example, access to functionality provided by external libraries that do not have a native VBA interface. Many libraries, including SQLite, expose the C-language API as exported DLL routines, possibly accessible to VBA code.

Modern Windows distributions ship with the SQLite library, though it may be a feature-limited build and not very recent. The most flexible approach is to [build][SQLite-ICU-MinGW] the current SQLite release from the source (the 32-bit version requires special [considerations][SQLite-Build-VBA]). The resulting single DLL file is usable even if placed in a user directory (additional DLL files may be necessary for the ICU extension).

There are several restrictions on VBA DLL calls. Usually, the [Declare][] statement must introduce the library's routines to the VBA compiler. It may also be necessary to load DLL explicitly via the Windows API if the library does not reside in a standard system directory. The bitness of the library host application must match. Finally, x32 libraries must follow the [STDCALL][]/WINAPI [ABI][] calling [convention][Calling convention].


<!-- References -->

[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[RDVBA]: https://rubberduckvba.com/
[RDBMS]: https://en.wikipedia.org/wiki/Relational_database
[VBA RDBMS]: https://bettersolutions.com/vba/databases/
[ADODB]: https://docs.microsoft.com/en-us/sql/ado/microsoft-activex-data-objects-ado
[SecureADODB]: https://github.com/rubberduck-vba/examples/tree/master/SecureADODB
[VBA SQL Library]: https://github.com/Beakerboy/VBA-SQL-Library
[vbaMyAdmin]: https://github.com/sauternic/vbaMyAdmin
[winsqlite3.dll-4-VBA]: https://renenyffenegger.ch/notes/development/databases/SQLite/VBA/
[SQLite for Excel]: https://github.com/govert/SQLiteForExcel/
[ChW SQLiteODBC]: http://www.ch-werner.de/sqliteodbc/
[SQLiteODBC PG]: https://pchemguy.github.io/SQLite-ICU-MinGW/odbc
[SQLite-ICU-MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
[SQLite-Build-VBA]: https://pchemguy.github.io/SQLite-ICU-MinGW/stdcall
[SQLite]: https://sqlite.org/
[Declare]: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/declare-statement
[STDCALL]: https://docs.microsoft.com/en-us/cpp/cpp/argument-passing-and-naming-conventions
[ABI]: https://en.wikipedia.org/wiki/Application_binary_interface
[Calling convention]: https://en.wikipedia.org/wiki/X86_calling_conventions
