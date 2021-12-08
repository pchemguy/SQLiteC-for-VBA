I started this project as the [SQLiteDB VBA][] class, wrapping the ADODB library to facilitate introspection of the SQLite engine and databases. Later, I refactored the SQLiteDB class and several supporting class modules into the SQLiteADO subpackage shown on the left in [Fig. 1](#LibraryStructure). SQLiteADO incorporates a set of class modules with a shared prefix *Lite-*. Shown on the right, the other core subpackage SQLiteC uses SQLite C-API directly (and the *SQLiteC-* prefix).

<a name="LibraryStructure"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/Major%20Componenets.svg" alt="Library structure" /></div>
<p align="center"><b>Fig. 1. Library structure</b></p>  

While I significantly refactored the code and added some features, the package is still focused on

* database connectivity
	* SQLiteADO - ADO/SQLiteODBC connection string helpers and limited ADODB wrappers,
	* SQLiteC - alternative connectivity approach bypassing ADO/SQLiteODBC,
* validation/integrity, and
* metadata (SQL-based SQLite introspection).


See [docs][] for further details.


<!-- References -->

[SQLiteDB VBA]: https://pchemguy.github.io/SQLiteDB-VBA-Library/
[docs]: https://pchemguy.github.io/SQLiteC-for-VBA/
