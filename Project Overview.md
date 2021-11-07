---
layout: default
title: Project overview
nav_order: 2
permalink: /project-overview
---

### Major library components

I started this project as the [SQLiteDB VBA][] class, wrapping the ADODB library to facilitate introspection of the SQLite engine and databases. Later, I refactored the SQLiteDB class and several supporting class modules into the SQLiteADO subpackage shown on the left in [Fig. 1](#LibraryStructure). SQLiteADO incorporates a set of class modules with a shared prefix *Lite-*. Shown on the right, the other core subpackage SQLiteC uses SQLite C-API directly (and the *SQLiteC-* prefix).

<a name="LibraryStructure"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/Major%20Componenets.svg" alt="Library structure" width="80%" /></div>
<p align="center"><b>Fig. 1. Library structure</b></p>  

The ILiteADO interface module shown in the center is a part of the SQLiteADO subpackage. It formalizes the high-level core functionality necessary for database interaction. As an OOP exercise, I implemented ILiteADO by the SQLiteC subpackage as well. An application may interact with SQLiteADO/SQLiteC directly via their APIs, which differ substantially. Or it may use the ILiteADO interface, which only exposes partial functionality but provides a unified API.

### Development environment

I use Excel 2002 (x32/VBA6) as my primary development environment. Additionally, I run tests under Excel 2016 (x64, VBA7). *SQLite C/ADO with Introspection for VBA* lives within the *SQLiteCAdoReflectVBAdev.xls* Excel Workbook located in the repository root. [Rubber Duck VBA][Rubber Duck VBA] VBA IDE add-in has become an essential VBA development component for me. I also regularly use [RDVBA Project Utils][] for exporting/importing the virtual VBA project structure.

### VBA project structure

I assumed the following convention for structuring my project, as seen from the Code Explorer. First of all, in the root of the virtual directory structure, I have the *Common* directory for reusable components, such as *RDVBA Project Utils*. For example, I usually have these components placed in subfolders under *Common*:

* *CPearson/Array* contains the [Chip Pearson's Array library][CPearson Array] with some fixes,
* *Guard* contains the Guard component used for input validation and testing,
* *Project Utils* contains the *RDVBA Project Utils* package, and
* *Shared* contains several regular modules with various helper routines.

There are three other root-level directories in this project. The *SQLite* directory contains sufficiently mature components, which are the main focus of this project. *SQLiteDBdev* includes various experimental/draft modules. Finally, the *DllManager* directory contains the *DLL Manager* developed specifically for the SQLiteC subpackage, but it may be useful in other projects and is described later in this documentation.

I do not use the top-level directory *Tests*. Instead, I prefer having individual test modules as close to their targets as possible. In this relatively large project, I decided to make one exception. While I usually place test fixtures within test modules, in this case, it would mean a lot of duplicated code. Further, some of them are complex enough to justify tests of their own. For these reasons, I grouped test fixtures into several separate *Fix-* prefixed modules. To avoid global namespace pollution, I decided to use predeclared classes rather than standard modules. And because these modules serve the entire project, they essentially constitute a test fixtures subpackage. I added tests to this subpackage and placed it under the main project directory (*SQLite*).

### Repository structure

The main project file hosting this VBA project is the *SQLiteCAdoReflectVBAdev.xls* Excel Workbook located in the repository root. The repo also contains three directories. The *Project* directory contains all project code modules exported by running the *ProjectUtilsSnippets.ProjectFilesExport* macro from the *Common/Project Utils* virtual project folder. The *Assets* directory hosts documentation figures. Finally, the *Library* directory is for package-specific files.

Each package should host its files in a subdirectory under *Library*. This way, packages with all their supporting files can be added to other projects easily. For example, the *DllManager* subdirectory contains custom-build SQLite binaries (both x32 and x64). Several DLL Manager tests use these binaries as test fixtures. Similarly, *SQLiteCAdo* holds a demo database and various files used by tests, demos, and examples.


<!-- References -->

[Rubber Duck VBA]: https://rubberduckvba.com
[RDVBA Project Utils]: https://pchemguy.github.io/RDVBA-Project-Utils/
[SQLiteDB VBA]: https://pchemguy.github.io/SQLiteDB-VBA-Library/
[CPearson Array]: http://www.cpearson.com/Excel/VBAArrays.htm
