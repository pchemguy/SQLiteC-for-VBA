---
layout: default
title: Customize & Extend DLL
nav_order: 2
parent: DLL Manager
permalink: /dllman/dll-custom-ext
---

This section discusses a couple of templates demonstrating the development of custom(ized) DLLs for use with VBA. The first example builds a minimalistic demo DLL from scratch and the other demos extension of the SQLite code.

**AddLib - Custom DLL from scratch**

AddLib demo follows this [tutorial][Transmission Zero], and its [folder][AddLib] includes the following files:

1. "add\.c"/"add\.h" - the C sources exporting one function and two variables;
2. "AddLib\.sh" - bash script building the library (to be run from a [MinGW][] shell);
3. "x32\\AddLib\.dll" - compiled 32-bit binary (x64 version has not been compiled/tested yet).

AddLibDemo.bas module in the "SQLiteC.xls" file (in the repo root) contains the VBA code loading and calling this demo library (for now, only the 32-bit version is available).

---

**SQLite - Extending DLL**

Given the source code of the DLL library, there are several approaches to extend it.

1. Patching the source code is optimal for small modifications/additions and should provide the best compile-time code optimization. As illustrated by this demo, a library's code module may contain the new code inlined, or it may #include the extension source files.  
2. Static linking reduces source code coupling. In this case, the extension modules are compiled into separate object modules statically linked with the library into a single binary file. For example, The [SQLite3odbc.dll][SQLiteODBC] driver (also see my [customized version][SQLiteODBC PG]) incorporates the SQLite library and the ODBC driver functionality.
3. Dynamic linking minimizes the coupling between the library and extension. The latter is compiled into a separate DLL dynamically linked to the library. [SQLiteForExcel][] project follows this approach. The stubs, performing STDCALL&nbsp;&rarr;&nbsp;CDECL translation for the SQLite API exposed by an official SQLite3 binary (x32), form the "SQLite3_StdCall.dll" module dynamically linked to the "SQLite3.dll" library module.

This demo consists of a single Bash script, [sqlite3\.ref\.extends\.sh][Dll adapter], based on the [code][SQLite ICU MinGW - Proxy] from a related [project][SQLite ICU MinGW]. The script compiles the library using the default CDECL convention. For this reason, an attempt to call any API from 32-bit VBA should cause the "Bad calling convention" error unless such API takes no arguments (see sqlite3_libversion_number call in the *DllExtAdapterEmbedDemo.GetSQLiteVersion* sub). At the same time, a new routine, *demo_sqlite3_extension_adapter*, is added to the SQLite3.c source code file before compilation. This adapter is labeled with STDCALL and, therefore, accessible from VBA.

---

There are a few other resources related to DLL development, which I added to my bookmarks: [Dynamic-Link Libraries][] and [C/C++ projects and build systems in Visual Studio][] from Microsoft and [DLL Tutorial][] from Tutorials Point.

<!-- References -->

[Transmission Zero]: https://www.transmissionzero.co.uk/computing/advanced-mingw-dll-topics/
[AddLib]: https://github.com/pchemguy/SQLiteC-for-VBA/tree/develop/Library/SQLiteCforVBA/Demo%20-%20DLL%20-%20STDCALL%20and%20Adapter/AddLib
[MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/devenv
[SQLite ICU MinGW]: https://pchemguy.github.io/SQLite-ICU-MinGW/
[SQLite ICU MinGW - Proxy]: https://github.com/pchemguy/SQLite-ICU-MinGW/blob/master/MinGW/Proxy/sqlite3.ref.sh
[SQLiteODBC]: http://www.ch-werner.de/sqliteodbc/
[SQLiteODBC PG]: https://pchemguy.github.io/SQLite-ICU-MinGW/odbc
[SQLiteForExcel]: https://github.com/govert/SQLiteForExcel
[Dll adapter]: https://github.com/pchemguy/SQLiteC-for-VBA/tree/develop/Library/SQLiteCforVBA/Demo%20-%20DLL%20-%20STDCALL%20and%20Adapter/SQLite
[Dynamic-Link Libraries]: https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-libraries
[C/C++ projects and build systems in Visual Studio]: https://docs.microsoft.com/en-us/cpp/build/projects-and-build-systems-cpp
[DLL Tutorial]: https://www.tutorialspoint.com/dll/
