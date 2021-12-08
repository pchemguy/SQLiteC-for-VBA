/* addclient.c

   Demonstrates using the function imported from the DLL, in a flexible and
   elegant way.
*/

#include <stdlib.h>
#include <stdio.h>
#include "add.h"

int main(int argc, char** argv)
{
  printf("6 + 23 = %d\n", Add(6, 23));

  return EXIT_SUCCESS;
}
