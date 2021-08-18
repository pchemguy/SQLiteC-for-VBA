SQLiteDB VBA library is a set of VBA helper functions for the SQLite engine. The primary motivation behind this project is to provide convenient access to extended introspection features. Because of its generality, ADODB/ADOX libraries provide only limited metadata information. Furthermore, this information is only available if the underlying driver implements the relevant functionality. The *Introspection* subpackage of this library, on the other hand, relies on the generic SQL querying mechanism and specialized SQL queries. It facilitates access to complete information about both the features of the active engine used and objects/attributes of the attached database.

The benefits of examining the metadata are twofold. For once, SQLite employs a modular architecture, with a significant portion of its widely used functionality provided via extensions. The availability of this functionality depends on the used compilation options, and some extensions, such as the ICU Unicode extension, also require external libraries. Thus, an application relying on the system copy of SQLite should verify that the necessary features are available. On the other hand, analyzing the structure of the attached database may add coding flexibility. This information may reduce the necessity for hardcoding specific database objects or provide a means to catch issues early (for example, when a third-party application modifies the database incorrectly).

The SQLiteDB VBA library uses the ADODB package and relies on the Christian Werner's [SQLiteODBC][SQLiteODBC CW] driver. (Please note that the driver has not been updated for over a year. While it should work, I use a custom compiled binary that embeds a recent version of SQLite, as described [here][SQLiteODBC PG].) My primary development environment is Excel XP/2002 x32, but I also run the test suite on Excel 2016 x64. The main file containing the current version of the library is *SQLiteDBVBALibrary.xls*. The *Project* folder contains all code modules exported using [RDVBA Project Utils][].

<a name="FigClassDiagram"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/SQLiteDB-VBA-Library/develop/Assets/Diagrams/Class Diagram.svg" alt="Class Diagram" width="100%" />  
<p align="center"><b>Figure 1. Class diagram</b></p>  

From the calling code's perspective, the top-level API object is *SQLiteDB* (Fig. 1). It uses the *Introspection* subpackage to generate appropriate SQL code, runs the query via the ADODB library, and exposes the resulting ADODB.Recordset object.

The main class of the *Introspection* subpackage, *SQLiteSQLDbInfo*, contains most of the routines generating SQL code used to obtain database-related information. A separate module *SQLiteSQLDbIdxFK* is responsible for bulky code related to indices and foreign keys, and the main subpackage class exposes this functionality via proxies. SQL code focused on the engine-related information is provided by the *SQLiteSQLEngineInfo* class exposed on the main subpackage class via encapsulation.

*SQLlib* module is responsible for generic SQL code, and *ADOlib* contains routines manipulating ADODB objects and related helper routines. *VerifyOrGetDefaultPath* routine from *ShellRoutines.bas* resolves database filepath with several checks and fallbacks, as described in the code comments.

I plan to add further details later; for the time being, please refer to the comments in the code, the test suite, and modules with examples. For example, *SQLiteIntropectionExample.bas* contains calls to all introspection routines, and the individual responses are placed on the corresponding Excel worksheets via the QueryTable feature.



<!-- References -->

[SQLiteODBC CW]: http://www.ch-werner.de/sqliteodbc/
[SQLiteODBC PG]: https://pchemguy.github.io/SQLite-ICU-MinGW/odbc
[RDVBA Project Utils]: https://pchemguy.github.io/RDVBA-Project-Utils/