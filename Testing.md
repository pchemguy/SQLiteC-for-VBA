---
layout: default
title: Testing
nav_order: 9
permalink: /testing
---

This project relies on the RDVBA testing framework, making the RubberDuck Add-in a required testing dependency (in fact, because of the test modules, the project will fail to compile/run without this add-in). Usually, I place the necessary fixture within the test modules before the test cases. This project follows a different approach, grouping the fixtures into several dedicated predeclared class modules ('Fix\*' modules located in the *SQLite/Fixtures* folder in RDVBA Code Explorer). These classes generate test objects and SQL snippets via their predeclared instances.

*FixObjC* and *FixObjAdo* class modules include factories for class instances used by various tests. For example, they generate connection objects associated with

* existing database,
* new in-memory database (both empty and populated with test data),
* new file-based database in the temp folder (same as above).

Except for some specialized cases, the file-based and in-memory databases should behave identically, so factories generating in-memory databases are preferred. Some tests use file-based databases due to particular needs. Occasionally, file-based databases are selected to enable manual examination of the generated databases.

*FixSQLBase*, *FixSQLFunc*, and *FixSQLITRB* classes provide SQL query templates used by tests and test factories. The SQLite introspection query returning the table of available functions is a convenient source of test data used by the *FixSQLFunc* module facilitating 'function table' based tests. Similarly, the *FixSQLITRB* module focuses on a small synthetic test dataset, including all basic data types.

*FixUtils* module provides a few generic convenience routines.

The *Library/SQLiteCAdo* and *Library/SQLiteCAdo/Fixtures* folders of the repository include several test databases populated with mock data and several database files, which were corrupted in a particular fashion to fail specific [integrity/validation][] checks. The two batch scripts in the latter folder set/reset special access permissions necessary for the *ztcCreate_FailsOnLastFolderACLLock* test, which yields an inconclusive result if the *acl-restrict.bat* script is not executed before testing. It is prudent to restore permissions by running *acl-restore.bat* to avoid possible issues with file system access.


<!-- References -->

[integrity/validation]: ./integrity
