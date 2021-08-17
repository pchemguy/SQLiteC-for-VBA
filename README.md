### Overview

SQLiteDB VBA Library is a set of VBA helper functions for the SQLite library. It uses ADODB connectivity and relies on the Christian Werner's [SQLiteODBC][SQLiteODBC CW] driver. (Please note that the driver has not been updated for a while. While it should work, I use custom compiled binaries, including a recent version of SQLite, as described [here][SQLiteODBC PG].) My primary development environment is Excel XP/2002 x32, but I also run the test suite on Excel 2016 x64.

The main file containing the current version of the library is *SQLiteDBVBALibrary.xls*. The *Project* folder contains all code modules exported using [RDVBA Project Utils][].

I plan to document this project properly later; for the time being, please refer to the comments in the code and the test suite for further details.



<-- References -->

[SQLiteODBC CW]: http://www.ch-werner.de/sqliteodbc/
[SQLiteODBC PG]: https://pchemguy.github.io/SQLite-ICU-MinGW/odbc
[RDVBA Project Utils]: https://pchemguy.github.io/RDVBA-Project-Utils/
