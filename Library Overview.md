---
layout: default
title: Library overview
nav_order: 1
permalink: /
---

SQLiteCAdo is an object-oriented database library compatible with Excel 2002/VBA6-x32 and 2016/VBA7-x64. It provides two alternative options for connecting to SQLite databases from VBA. A common approach relies on the ADODB library and the SQLiteODBC driver. The latter, however, must be installed on the target computer by an administrator. As of this writing, the latest stable release of the driver, dated June 2020, embeds a feature-limited, outdated copy of SQLite.

Talking to the SQLite library directly via its C-language API provides an alternative approach. SQLiteCAdo dynamically loads an SQLite DLL from the project directory via Windows API and does not require installation/registration of the DLL. The project includes SQLite binaries for Windows 10 custom-built from a recent release with all features activated, including case insensitive operations for non-ASCII characters via ICU. The repository contains x32 and x64 binaries, building guidelines, and associated scripts. 

An application incorporates the SQLiteCAdo library via its source code and accesses it via its objects. Because of efficiency considerations, the library follows the data mapper pattern and does not provide any ORM-like features. After initialization, it takes SQL queries as input and transfers the data between the database and VBA data structures (Recordsets and 2D arrays). SQLiteCAdo also incorporates a limited suite of SQL helpers emphasizing SQL-based introspection.

The to-do list includes several items. The ADODB wrapper does not handle parameterized queries yet. With the C-API wrapper, the bulk UPDATE operation presently requires looping through the records in the application. Another thin wrapper may take care of bulk updates while facilitating integration with the Model-View-Presenter (MVP) pattern. The C-API wrapper may also yield more efficient database operations, but whether it does remains to be seen.

<a name="LibraryStructure"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/Major%20Componenets.svg" alt="Library structure" /></div>
<p align="center"><b>Fig. 1. Library structure</b></p>  
