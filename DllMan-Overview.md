---
layout: default
title: Overview
nav_order: 1
parent: DLL Manager
permalink: /dllman/overview
---

### VBA class managing non-system DLL libraries

I developed this class as a standalone component necessary for the SQLiteC package. With SQLiteC, I wanted to use custom-built SQLite binaries placed within the project directory. Before a DLL library placed in a user directory is available from the VBA code, it must be loaded via the Windows API (alternatively, the Declare statement may include the library location, but this approach is ugly and inconvenient). It is also prudent to unload the library when no longer needed. To make the load/unload process more robust, I created the DllManager class wrapping the LoadLibrary, FreeLibrary, and SetDllDirectory [APIs][DLL API]. DllManager can be used for loading/unloading multiple DLLs. It wraps a Scripting.Dictionary object to hold \<DLL name\>&nbsp;&rarr;&nbsp;\<DLL handle\> mapping.

**API**

The *DllManager.Create* factory takes three optional parameters, passed to the *DllManager.Init* constructor:

*	*DefaultPath* (defaults to blank) - a string indicating the default DLL directory,
*	*DllNames* (defaults to empty) - a string containing the DLL name to be loaded or a variant array of DLL names,
*	*GetShared* (defaults to true) - a boolean indicating whether the factory should return the singleton instance.

The *DefaultPath* setter (Property Let) handles *DefaultPath*. The setter checks if the parameter holds a valid absolute or a relative (w.r.t. ThisWorkbook.Path) path. If this check succeeds, SetDllDirectory API sets the default DLL search path. *DllManager.ResetDllSearchPath* can be used to reset the DLL search path to its default value. If the singleton object is ever requested, the predeclared DllManager will instantiate a new object on the first request, save its reference, and return it in response to successive requests. A new independent instance can still be requested at any time by setting the third parameter to false. *DllManager.ForgetSingleton* clears saved singleton reference (this method should be called on the predeclared instance).

*DllManager.Load* loads individual libraries. It takes the target library name and, optionally, path. If the target library has not been loaded, it attempts to resolve the DLL location by checking the provided value and the DefaultPath attribute. If resolution succeeds, the LoadLibrary API is called. *DllManager.Free*, in turn, unloads the previously loaded library.

*DllManager.LoadMultiple* loads a list of libraries. It takes a variable list of arguments (ParamArray) and loads them in the order provided. Alternatively, it also accepts a 0-based array of names as the sole argument. *DllManager.FreeMultiple* is the counterpart of *.LoadMultiple* with the same interface. If no arguments are provided, all loaded libraries are unloaded.

Finally, while *.Free/.FreeMultiple* can be called explicitly, *Class_Terminate* calls  *.FreeMultiple* and *.ResetDllSearchPath* automatically before the object is destroyed.

**Demo**

The *DllManagerDemo* example below illustrates how this class can be used and compares the usage patterns between system and user libraries. In this case, *WinSQLite3* system library is used as a reference (see *GetWinSQLite3VersionNumber*). A call to a custom compiled SQLite library placed in the project folder demos the additional code necessary to make such a call (see *GetSQLite3VersionNumber*). In both cases, *sqlite3_libversion_number* routine, returning the numeric library version, is declared and called.


<!-- References -->

[DLL API]: https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-functions
[SQLite VBA]: https://pchemguy.github.io/SQLite-ICU-MinGW/stdcall
