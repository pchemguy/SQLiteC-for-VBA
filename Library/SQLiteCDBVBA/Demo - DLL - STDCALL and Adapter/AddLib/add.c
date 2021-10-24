#include "add.h"


ADDAPI int ADDCALL Add(int a, int b) {
  return (a + b);
}

/* Assign value to exported variables. */
ADDAPI int foo = 7;
ADDAPI int bar = 41;
