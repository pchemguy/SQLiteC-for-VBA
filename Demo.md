---
layout: default
title: Demos and examples
nav_order: 3
permalink: /demo
---

Dedicated demo code is located in several modules:

* **LiteExamples** in SQLite/ADO/ADemo
* **SQLiteCExamples** in SQLite/C/ADemo
* **SQLiteCAdoDemo** in SQLite/AADemo

Additionally, **FixObj*** in SQLite/Fixtures and test modules provide further usage examples.

**LiteExamples** module contains several small basic independent routines, which do not demonstrate typical database commands (see below).

**SQLiteCExamples** module provides usage examples for various SQLiteC objects, including database operations. The Main Sub is the primary entry point running the demo.

**SQLiteCAdoDemo** compares the use of SQLiteC (MainC) and SQLiteADO (MainADO) subpackages via the ILiteADO interface. The first 3-4 lines at the top of the Main* modules perform initialization of the ILiteADO **dbq** object (SQLiteC needs a bit more initialization code, which is placed in a separate Sub with debug messages) and create a new blank SQLite database file in the Temp folder.  Both routines call the DemoDBQ Sub after initialization is complete. This Sub demonstrates typical operations: it creates two demo tables and illustrates insert, select, and update statements. The two conditional structures show differences in implemented features. SQLiteC supports parameterized queries, including via the ILiteADO interface, while SQLiteADO does not support them.
