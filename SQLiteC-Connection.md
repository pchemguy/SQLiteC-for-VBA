---
layout: default
title: Connection
nav_order: 2
parent: SQLiteC
permalink: /sqlitec/connection
---

Functionally, the SQLiteCConnection class is roughly modeled after the ADODB.Connection class. As a result, however, SQLiteCConnection is a fairly large class (a candidate for refactoring/splitting). It is usually instantiated via the SQLiteC.CreateConnection method and is responsible for several groups of functions.

* Database connection
   OpenDb and CloseDb methods wrap APIs responsible for connection handling.
* Error information
   ErrInfo, GetErr, PrintErr, and ErrInfoRetrieve methods wrap APIs retrieving error information.
* Connection information
   AttachedDbPathName and AccessMode provide file pathname and access mode for a specified attached database;
   ChangesCount returns the number of changes for the last modifying transaction or the total number of changes performed via specific connection.
* Plain queries
   ExecuteNonQueryPlain wraps the *exec* API. The SQLite engine provides a large set of APIs necessary to execute a query and retrieve the result. It also provides a simplified "shortcut" interface *exec*. This interface does not support parametrized queries and in order to retrieve the result, a callback function must be used. At the same time, it handles muti-statement quries automatically, whereas the primary query interface stops processing the query command at the first semicolon. For simplicity reasons, ExecuteNonQueryPlain does not use a callback function and cannot return any data. It is useful for executing control queries not returning data.  
* Database manipulation
   Attach, Detach, and Vacuum provide corresponding functionality.
