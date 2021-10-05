/* You should define ADD_EXPORTS *only* when building the DLL. */
#ifdef ADD_EXPORTS
  #define ADDAPI __declspec(dllexport)
#else
  #define ADDAPI __declspec(dllimport)
#endif

/* Define calling convention in one place, for convenience. */
#define ADDCALL __stdcall

/* Make sure functions are exported with C linkage under C++ compilers. */
#ifdef __cplusplus
extern "C"
{
#endif

/* Declare our Add function using the above definitions. */
//__declspec(dllexport) __stdcall int Add(int a, int b);
int Add(int a, int b);

/* Exported variables. */
extern int foo;
extern int bar;

#ifdef __cplusplus
} // __cplusplus defined.
#endif