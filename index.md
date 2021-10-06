---
layout: default
title: Introduction
nav_order: 1
permalink: /
---

First introduced for automating MS Excel, VBA is still the only native embedded programming language available in MS Office. VBA is an object-oriented scripting language requiring no additional software for development (though [RDVBA][] is a great, if not essential, addition). MS Excel, in turn, provides a convenient means for data management, manipulation, and presentation, requiring little to no programming experience and development time, and is, thus, appealing to a broad spectrum of users. A wide variety of data-focused tasks, both personal and business-related, take advantage of Excel functionality even when an [RDBMS][] database would do a better job.

VBA incorporates several [technologies][VBA RDBMS] for accessing databases, providing a means for Excel to interact with database engines. These database-related VBA libraries present various levels of abstraction and, depending on the database, may require additional drivers to be installed. Through these libraries, VBA can interact with many database engines, and for personal applications, SQLite is likely a good choice.

SQLite is notable for its small size and has a simplified architecture that does not involve a database server. VBA provides access to SQLite databases via the [ADODB][] library and the [SQLiteODBC][] driver. The driver embeds a copy of the SQLite library, so no other software needs to be installed. The driver has a few limitations, however. As of this writing (Fall 2021), the driver has not been updated for more than a year, and it incorporates an outdated copy of SQLite.  Furthermore, the official SQLiteODBC binaries do not include valuable SQLite functionality developed as optional extensions. Some of these issues can be resolved if the driver is [custom-built][SQLiteODBC PG] from the source.

Alternatively, VBA can access SQLite databases directly, bypassing both the ADO library and the SQLiteODBC driver. VBA does not need any database drivers for direct access to SQLite, and the only required component is the SQLite library file. Modern Windows distributions ship with the SQLite library, though it may be feature-limited and not very recent. The most flexible approach is to [build][SQLite-ICU-MinGW] the current SQLite release from the source (the 32-bit version requires special [considerations][SQLite-Build-VBA]). The resulting single dll file can be used even if placed in a user directory (additional dll files will be necessary if the ICU extension is enabled).

Direct access to SQLite relies on the VBA capability to call dll routines directly. This feature can be used, for example, to access functionality provided by external libraries that do not have a native VBA interface. Many libraries expose the C-language API as exported DLL routines, and by calling these routines, VBA code may be able to use such APIs. SQLite, written in the C-language, is an example of such a library. It is commonly used as a dynamically loaded library and naturally provides the C-language API.

Before the dll library's functionality is available from within the VBA code, the library's routines must be introduced via the [Declare][] statement. Additionally, if the called library does not reside in a standard system directory, it may also be necessary to load such a library explicitly via the Windows API. The bitness of the called library must match the bitness of the VBA host application, and for 32-bit hosts, routines called by VBA must follow the [STDCALL][]/WINAPI [ABI][] calling [convention][Calling convention].

<!-- References -->

[RDVBA]: https://rubberduckvba.com/
[RDBMS]: https://en.wikipedia.org/wiki/Relational_database
[VBA RDBMS]: https://bettersolutions.com/vba/databases/
[ADODB]: https://docs.microsoft.com/en-us/sql/ado/microsoft-activex-data-objects-ado
[SQLiteODBC]: http://www.ch-werner.de/sqliteodbc/
[SQLiteODBC PG]: https://pchemguy.github.io/SQLite-ICU-MinGW/odbc
[SQLite-ICU-MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
[SQLite-Build-VBA]: https://pchemguy.github.io/SQLite-ICU-MinGW/stdcall
[SQLite]: https://sqlite.org/
[Declare]: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/declare-statement
[STDCALL]: https://docs.microsoft.com/en-us/cpp/cpp/argument-passing-and-naming-conventions
[ABI]: https://en.wikipedia.org/wiki/Application_binary_interface
[Calling convention]: https://en.wikipedia.org/wiki/X86_calling_conventions
