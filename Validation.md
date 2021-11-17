---
layout: default
title: Validation and integrity
nav_order: 6
permalink: /integrity
---

### Integrity and validation: LiteFSCheck, LiteACID, and CommonRoutines

LiteFSCheck and LiteACID classes (SQLite/Checks) with the CommonRoutines module (Common/Shared) validate connection settings and consistency/integrity related functionality.

### LiteFSCheck

The sole function of the LiteFSCheck class is to validate/resolve database specifications and perform file-system-based checks if appropriate. LiteFSCheck is a predeclared class, and its factory (default member) has the same signature as that of LiteMan. If a file path is provided, the second parameter controls whether the specified file must be a valid existing database file. When opening an existing file, the second argument must be set explicitly to *False*. Apart from relative/absolute file path, the LiteFSCheck factory also accepts these names as the first argument:

 * ":memory:" or ":mem:" for an in-memory database,
 * ":temp:" or ":tmp:" for a new file-based database located in the Temp folder (with file name generated using date, time, and a random suffix string),
 * ":blank:" for an anonymous file-based database.

LiteFSCheck contains three functional sections: one focused on stepwise verification of the path, another block, if appropriate, executes a series of checks aiming at verifying that the file is accessible and appears to be a valid SQLite3 database file, and the third (entry routine) is a dispatcher. The dispatcher handles shortcut names described above and passes control to the first two routines as necessary. For an existing file, it will also call the path resolution function VerifyOrGetDefaultPath in CommonRoutines. LiteFSCheck factory calls its constructor, which, in turn, runs all tests according to provided arguments, so the instance of LiteFSCheck returned by the factory already has the check results available.

<p align="center"><b>Table 1. Sample immediate pane commands</b></p>

| Command                                                 | Result                       |  
|---------------------------------------------------------|------------------------------|  
| `?LiteFSCheck(":mem:").DatabasePathName`                | ":memory:"                   |  
| `?LiteFSCheck(":tmp:").DatabasePathName`                | Path to a new db in Temp     |  
| `?LiteFSCheck("SQLiteCAdo.db", False).DatabasePathName` | Path to an existing database |  

### LiteACID

SQLite provides three commands for checking the database integrity:
 
 * [pragma_integrity_check][],
 * [pragma_quick_check][] (a quicker version of the former, running a subset of checks),  and
 * [pragma_foreign_key_check][].

LiteACID follows the same pattern as other classes, having the predeclared attribute and Create factory as the default member. The factory takes an ILiteADO instance and returns a LiteACID instance. LiteACID's primary method, IntegrityADODB, runs a full integrity check followed by the foreign-key check and returns True if both tests pass. Otherwise, it prints out a message, indicating which test failed. Usually, it is not necessary to save the LiteACID instance. For one-time use, the factory call accepts a chained call to IntegrityADODB, e.g., the following *immediate pane* command

    ?LiteACID(LiteADO(LiteFSCheck("SQLiteCAdo.db", False).DatabasePathName)).IntegrityADODB

should produce output

    -- Integrity check passed for: <full pathname>
    True

LiteMan and LiteACID have several overlapping methods related to the journal mode feature. This minor defect should be fixed in the future.

<!-- References -->

[pragma_quick_check]: https://www.sqlite.org/pragma.html#pragma_quick_check
[pragma_integrity_check]: https://www.sqlite.org/pragma.html#pragma_integrity_check
[pragma_foreign_key_check]: https://www.sqlite.org/pragma.html#pragma_foreign_key_check
