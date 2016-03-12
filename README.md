# Overview
[SQLite](http://www.sqlite.org) is a small, easy-to-use, open-source SQL database engine. This project, *SQLite for Excel*, is a lightweight wrapper to give access to the SQLite3 library from VBA. It provides a high-performance path to the SQLite3 API functions, preserving the semantics of the SQLite3 library calls and allowing access to the distributed SQLite3.dll without recompilation.

The current release has the following parts:
* *SQLite3_StdCall* is a small and very simple C .dll that makes it possible to use the standard SQLite3 .dll from VBA. It just passes calls from VBA on to SQLite without any change in the parameters, but this allows the StdCall calling convention that VB6 and VBA is limited to.
* *Sqlite3.bas VBA module* has all the VBA Declares, and does the parameter and string conversions. It exposes a number of SQLite3xxxx functions. These map as directly as possible to the SQLite C API, with no change in the semantics. Although I have not exposed the whole API, most of the core interface is included, in particular the prepared statement, binding, retrieval and backup functions. Date values are stored as Julian day real numbers in the database.
* *Sqlite3Demo.bas VBA module* has tests that serve as nice examples of how to use the SQLite3xxxx functions. 
* *SQLite3Demo.xls* contains the two VBA modules.

* *64-bit support* for use with the 64-bit versions of Excel can be found in {"SQLiteForExcel_64.xlsm"} which has VBA code that supports both 32-bit and 64-bit versions of Excel. A 64-bit build of SQLite 3.7.13 is located in x64\SQLite3.dll. The corresponding {"Sqlite3Demo_64.bas"} module shows how to target both 32-bit and 64-bit Excel with the same VBA code (some #Ifs are required). (Note that the default install of Office is always the 32-bit version, even on a 64-bit version of Windows. Only if the 64-bit version of Office has been specifically selected will the 64-bit modules be required.) 

# Getting Started
* Download the release archive .zip file from https://github.com/govert/SQLiteForExcel/releases.
* Unzip the download to a convenient location.
* Open the Distribution\SQLiteForExcel.xls file.
* Open the VBA Editor (Alt+F11).
* Note the SQLite3 module which contains the declarations and helper functions to access SQLite.
* Examine and run the example test code in the SQLite3Demo module.
* Find the documentation for the SQLite functions here: http://sqlite.org/cintro.html. The complete query language for SQLite is documented here: http://sqlite.org/lang.html.

# Sample projects
* Mark Camilleri has posted a sample project, *XLSQLite*, that provides a GUI interface to SQLite in Excel. The project can be downloaded from the Gatekeeper for Excel site: http://www.gatekeeperforexcel.com/other-freebies.html. _XLSQLite.xlam is an add-in that allows you to create and manipulate (both DDL and DML) SQLite databases from within Excel.  It offers a basic gui interface that allows you to perform basic tasks on your SQLite databases._ (Note that this example might not include the latest SQLite for Excel version, which has a fix for a string conversion bug.)

* *SQLite for Access* - Thomas Gewinnus has ported the SQLite-Interface to Access-VBA and added a small DAO-like-Layer (class SQLiteDatabase). See Module Test__SQLiteDatabase for samples. Download the .accdb file from https://s3.amazonaws.com/share.excel-dna.net/Beispiel.zip.

# Related Projects
* The *SQLite* home is at http://www.sqlite.org and the most recent version of the SQLite3.dll library can be found here http://www.sqlite.org/download.html.
* To create User-Defined Functions (UDFs) for Excel using C#, VB.NET or F#, have a look at my [Excel-DNA](http://https://github.com/Excel-DNA/ExcelDna) project. It provides free and easy integration of .NET with Excel.
* For access to SQLite from .NET I recommend:
    * the official [System.Data.SQLite](http://system.data.sqlite.org) is a full-featured ADO.NET driver with full Linq and Entity Framework support, or
    * the sweet-looking [sqlite-net](https://github.com/praeclarum/sqlite-net), a light-weight wrapper with attribute-based object-to-database mapping and some Linq support.

# Support
Create a new [GitHub Issue](https://github.com/govert/SQLiteForExcel/issues). You are also welcome to contact me directly at mailto:govert@icon.co.za with questions, comments or suggestions. 
