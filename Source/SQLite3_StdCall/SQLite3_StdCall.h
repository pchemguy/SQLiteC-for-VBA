// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the SQLITE3_STDCALL_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// SQLITE3_STDCALL_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef SQLITE3_STDCALL_EXPORTS
#define SQLITE3_STDCALL_API __declspec(dllexport) 
#else
#define SQLITE3_STDCALL_API __declspec(dllimport)
#endif

// TODO: Add declarations .... ?
