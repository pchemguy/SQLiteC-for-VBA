#include "add.h"

//__declspec(dllexport) __stdcall int Add(int a, int b)
int Add(int a, int b)
{
  return (a + b);
}

/* Assign value to exported variables. */
int foo = 7;
int bar = 41;