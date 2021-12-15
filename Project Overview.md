---
layout: default
title: Project overview
nav_order: 4
permalink: /project-overview
---

### Major library components

I started this project as the [SQLiteDB VBA][] class, wrapping the ADODB library to facilitate introspection of the SQLite engine and databases. Later, I refactored the SQLiteDB class and several supporting class modules into the SQLiteADO subpackage shown on the left in [Fig. 1](#LibraryStructure). SQLiteADO incorporates a set of class modules with a shared prefix *Lite-*. Shown on the right, the other core subpackage SQLiteC uses SQLite C-API directly (and the *SQLiteC-* prefix).

<a name="LibraryStructure"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/Major%20Componenets.svg" alt="Library structure" /></div>
<p align="center"><b>Fig. 1. Library structure</b></p>  

The ILiteADO interface module shown in the center is a part of the SQLiteADO subpackage. It formalizes the high-level core functionality necessary for database interaction. As an OOP exercise, I provided SQLiteC/ILiteADO implementation as well. An application may interact with SQLiteADO/SQLiteC via their APIs, which differ substantially. Another option is to use the ILiteADO interface, which exposes partial functionality but provides a unified API.

### Development environment

I use Excel 2002 (x32/VBA6) as my primary development environment. *SQLite C/ADO with Introspection for VBA* lives within the *SQLiteCAdoReflectVBAdev.xls* Excel Workbook located in the repository root. [Rubber Duck VBA][Rubber Duck VBA] VBA IDE add-in has become an essential VBA development component for me (and it installs just fine without the admin privileges on the user's account, which is a nice bonus). I also regularly use [RDVBA Project Utils][] for exporting/importing the virtual VBA project structure.

### Compatibility end testing

This project uses the RDVBA's unit testing framework as the primary means for testing, which means RDVBA add-in is required to run the tests. As of this writing, I use RDVBA 2.5.2.5871 (I had some issues with the latest following release v2.5.2.1 / build 2.5.2.5906, as I do with the currently used version, but I am waiting for now for the next release). Testing-wise, this build has a known very annoying GUI-related issue, rendering the testing framework barely usable when the number of tests grows above, say, 100. As a workaround, I disable the display of tests with unknown and successful result statuses. Another RDVBA testing issue and a workaround are discussed in an [SO Q&amp;A][].

I run tests under x32/VBA6 (Excel XP/2002 SP3) and x64/VBA7 (Excel 2016) environments and might also include a small set of tests that will run without RDVBA. The project, therefore, should be compatible with both x32/VBA6 and x64/VBA7. The primary source of compatibility concerns is the declarations of API routines, and it boils down to three keywords: *PtrSafe*, *LongPtr*, and *LongLong*. For portability, I use conditional compilation coupled primarily with the VBA7 constant. I have only a couple of instances of *LongLong*, and I added the WIN64 test within the VBA7 block for those cases, and I defined it as Currency within the VBA6 code. While not tested, this arrangement should make the code compatible with x32/VBA7, but it may not work under x64/VBA6.

### Required library references

* Microsoft ActiveX Data Objects 6.1 Library
* Microsoft Scripting Runtime 1.0
* Microsoft Visual Basic for Application Extensibility 5.3 (including approved programmatic access to the VBA project)
* Microsoft VBScript Regular Expressions 5.5
* Windows Script Host Object Model 1.0

While SQLiteCAdo does not access the VBA project, the third item is necessary for the [RDVBA Project Utils][] library (Common/Project Utils).

### VBA project structure

I assumed the following convention for structuring my project in the RDVBA Code Explorer. First of all, in the root of the virtual directory structure, I have the *Common* directory for reusable components, such as *RDVBA Project Utils*. For example, I usually have these components placed in subfolders under *Common*:

* *CPearson/Array* - the [Chip Pearson's Array library][CPearson Array] with some fixes,
* *Guard* - the Guard component used for input validation and testing,
* *Project Utils* - the [RDVBA Project Utils][] package, and
* *Shared* - several standard modules with various helper routines.

There are three other root-level directories in this project. The *SQLite* directory contains sufficiently mature components, which are the main focus of this project. *SQLiteDBdev* includes various experimental/draft modules. Finally, the *DllTools* directory contains [DLL Manager][DllTools] developed specifically for the SQLiteC subpackage, but it may be useful in other projects and is described later in this documentation.

I do not use the top-level directory *Tests*. Instead, I prefer having individual test modules as close to their targets as possible. In this relatively large project, I decided to make one exception. While I usually place test fixtures within test modules, in this case, it would mean a lot of duplicated code. Further, some of them are complex enough to be tested also. For these reasons, I grouped test fixtures into several separate *Fix-* prefixed modules. To avoid global namespace pollution, I decided to use predeclared classes rather than standard modules. And because these modules serve the entire project, they form a pseudo subpackage. I added tests to this subpackage and placed it under the top project directory (*SQLite*).

### Repository structure

The main application file hosting this VBA project is the *SQLiteCAdoReflectVBAdev.xls* Excel Workbook located in the repository root. The repo also contains three directories. The *Project* directory contains all project code modules exported by running the *ProjectUtilsSnippets.ProjectFilesExport* macro from the *Common/Project Utils* virtual project folder. The *Assets* directory hosts documentation figures. Finally, the *Library* directory is for package-specific files.

Each package should host its files in a subdirectory under *Library*, avoiding potential collisions and improving portability. For example, the *DllTools* subdirectory contains custom-build SQLite binaries (both x32 and x64). Several DLL Manager tests use these binaries as test fixtures. Similarly, *SQLiteCAdo* holds a demo database and various files used by tests, demos, and examples.

### Supporting tools

I prepare diagrams starting from the [yWorks yEd][] graph editor. I save originals in the native GraphML format and export them in the EPS format. Then I open EPS files in Adobe Illustrator CS6 and save them as SVGs (I also make jpg or png files at this point if necessary).

[VBADecompiler][] removes compiled VBA code from the host file, making it safer and more compact.

[TableGenerator][] assists with table markdown table creation.

[Grammarly][] service can significantly facilitate the writing process and help improve the linguistic quality of the texts.



<!-- References -->

[Rubber Duck VBA]: https://rubberduckvba.com
[RDVBA Project Utils]: https://pchemguy.github.io/RDVBA-Project-Utils/
[SQLiteDB VBA]: https://pchemguy.github.io/SQLiteDB-VBA-Library/
[CPearson Array]: http://www.cpearson.com/Excel/VBAArrays.htm
[RDVBA Project Utils]: https://pchemguy.github.io/RDVBA-Project-Utils/
[DllTools]: https://pchemguy.github.io/DllTools/
[SO Q&amp;A]: https://stackoverflow.com/questions/70098835/excel-hangs-at-exit-after-running-rdvba-tests
[yWorks yEd]: https://www.yworks.com/products/yed
[VBADecompiler]: http://orlando.mvps.org/VBADecompilerMore.asp
[TableGenerator]: https://www.tablesgenerator.com/
[Grammarly]: https://www.grammarly.com/
