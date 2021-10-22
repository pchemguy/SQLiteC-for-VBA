/*
**
*************************************************************************
**
*/
#ifndef ADD_H
#define ADD_H

#ifdef _WIN32

  /* You should define ADD_EXPORTS *only* when building the DLL. */
  #ifdef ADD_EXPORTS
    #define ADDAPI __declspec(dllexport)
  #else
    #define ADDAPI __declspec(dllimport)
  #endif

  /* Define calling convention in one place, for convenience. */
  #define ADDCALL __stdcall

#else /* _WIN32 not defined. */

  /* Define with no value on non-Windows OSes. */
  #define ADDAPI
  #define ADDCALL

#endif /* _WIN32 */


/* Make sure functions are exported with C linkage under C++ compilers. */
#ifdef __cplusplus
extern "C"
{
#endif

/* Declare our Add function using the above definitions. */
ADDAPI int ADDCALL Add(int a, int b);

/* Exported variables. */
ADDAPI extern int foo;
ADDAPI extern int bar;

#ifdef __cplusplus
} // __cplusplus defined.
#endif

#endif /* ADD_H */
