/*
**
*************************************************************************
**
*/
#include <stdlib.h>
#include <stdint.h> 
#include <math.h> 
#include <stdio.h>
#include <time.h> 
#include <string.h> 

#ifndef MEMTOOLS_H
#define MEMTOOLS_H

#ifdef _WIN32

  /* You should define MEMTOOLS_EXPORTS *only* when building the DLL. */
  #ifdef MEMTOOLS_EXPORTS
    #define MEMTOOLSAPI __declspec(dllexport)
  #else
    #define MEMTOOLSAPI __declspec(dllimport)
  #endif

  /* Define calling convention in one place, for convenience. */
  #define MEMTOOLSCALL __stdcall

#else /* _WIN32 not defined. */

  /* Define with no value on non-Windows OSes. */
  #define MEMTOOLSAPI
  #define MEMTOOLSCALL

#endif /* _WIN32 */


/* Make sure functions are exported with C linkage under C++ compilers. */
#ifdef __cplusplus
extern "C" {
#endif

/* Declare our function using the above definitions. */
MEMTOOLSAPI void MEMTOOLSCALL CopyMem(void* Destination, const void* Source, size_t Length);
MEMTOOLSAPI  int MEMTOOLSCALL PerfGauge(unsigned int ForCount);

MEMTOOLSAPI void MEMTOOLSCALL DummySub0Args();
MEMTOOLSAPI void MEMTOOLSCALL DummySub3Args(void*, const void*, size_t);
MEMTOOLSAPI  int MEMTOOLSCALL DummyFnc0Args();
MEMTOOLSAPI  int MEMTOOLSCALL DummyFnc3Args(void*, const void*, size_t);

#ifdef __cplusplus
} // extern "C"
#endif

#endif /* MEMTOOLS_H */
