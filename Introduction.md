---
layout: default
title: Introduction
nav_order: 1
permalink: /
---

### Motivation

First introduced for automating MS Excel, VBA is still the only native embedded programming language available in MS Office. VBA is an object-oriented scripting language requiring no additional software for development (though [RDVBA][] is a great, if not essential, addition). MS Excel, in turn, provides a convenient means for data management, manipulation, and presentation, requiring little to no programming experience and development time, and is, thus, appealing to a broad spectrum of users. A wide variety of data-focused tasks, both personal and business-related, take advantage of Excel functionality even when an [RDBMS][] database would do a better job.

VBA incorporates several [technologies][VBA RDBMS] for accessing databases, providing a means for Excel to interact with database engines. These database-related VBA libraries present various levels of abstraction and, depending on the database, may require additional drivers to be installed. The [ADODB][] library is a medium-level library designed to be used directly by VBA applications. However, in the common usage scenarios, the entire repertoire of technical features exposed by the library may not be necessary. Additional "middleware" packages may wrap and encapsulate unnecessary details, simplifying the client code and workflows.

### Existing database-centric VBA "middleware" projects

A few existing VBA projects develop database-related tools worth mentioning. Here are three generic projects:

* [SecureADODB][] is a thin wrapper around the ADODB library. SecureADODB is a generic (database-agnostic) fully object-oriented VBA library focused on database connectivity workflows. The project employs advanced OOP techniques and patterns, demonstrating VBA implementation details, and presents powerful testing strategies.
* [VBA SQL Library][] targets the generation of SQL queries, another common component of the database interaction workflow. I have not studied its code yet or tried to use it, so I cannot comment on its functionality.
* [vbaMyAdmin][] focuses on providing a GUI database administration functionality. I have only recently discovered this application, so no further comments here either.

And the following two projects specifically aim at SQLite access from VBA:

* [winsqlite3.dll-4-VBA][] demonstrates the use of the winsqlite3.dll library, a part of Windows distribution, and can be accessed directly from both VBA x32 and x64. While I managed to run the code from the project repository, it kept crashing Excel.
* [SQLite for Excel][] provides a standard module with wrappers around SQLite C-language API. Further, since VBA-x32 cannot call official SQLite3 binaries directly, the project also provides a compiled x32 DLL adaptor. It has both advantages and disadvantages, as further later.

### SQLite

SQLite is notable for its small size and has a simplified architecture that does not involve a database server. Therefore, SQLite is likely a good choice for personal applications and may fit well certain business applications.

A common means of accessing SQLite databases from VBA is via the ADODB library and the [ChW SQLiteODBC][] driver. The driver embeds a copy of the SQLite library, so no other software needs to be installed. However, the driver has a few limitations. As of this writing (Fall 2021), ChW has not updated the driver since June 2020, and the embedded copy of SQLite is an outdated feature-limited build. While ChW also provides an experimental version of SQLiteODBC binaries that should work with a local copy of the SQLite library, I could not make it work. The alternative options include using a [custom-built][SQLiteODBC PG] SQLiteODBC driver or bypassing it entirely. This project explores both of these options.

Bypassing SQLiteODBC/ADODB relies on the VBA ability to call DLL routines directly. This feature can be used, for example, to access functionality provided by external libraries that do not have a native VBA interface. Many libraries expose the C-language API as exported DLL routines, and by calling these routines, VBA code may be able to use such APIs. SQLite, written in the C-language, is an example of such a library. It is commonly used as a dynamically loaded library and naturally provides the C-language API.

Modern Windows distributions ship with the SQLite library, though it may be a feature-limited build and not very recent. The most flexible approach is to [build][SQLite-ICU-MinGW] the current SQLite release from the source (the 32-bit version requires special [considerations][SQLite-Build-VBA]). The resulting single DLL file can be used even if placed in a user directory (additional DLL files will be necessary if the ICU extension is enabled).

Before the DLL library's functionality is available from within the VBA code, the [Declare][] statement must introduce the library's routines. Additionally, if the called library does not reside in a standard system directory, it may also be necessary to load such a library explicitly via the Windows API. The bitness of the called library must match the bitness of the VBA host application, and for 32-bit hosts, routines called by VBA must follow the [STDCALL][]/WINAPI [ABI][] calling [convention][Calling convention].


<!-- References -->

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
