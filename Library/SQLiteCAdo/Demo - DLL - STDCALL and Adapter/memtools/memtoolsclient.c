#include "memtools.h"


int main(int argc, char** argv) { 
  struct timeb start, end;
  int diff;
  const double MSEC_IN_SEC = 1000.0;

  int dest;
 
  ftime(&start);
  for (int i=0; i < 1e8 ; i++) {
    CopyMem(&dest, &i, sizeof dest); 
  }
  ftime(&end);
  diff = MSEC_IN_SEC * (end.time - start.time)
                     + (end.millitm - start.millitm);

  printf("\nOperation took %u milliseconds\n", diff);
  printf("\nFinal value: %u\n", dest);
  return 0;
}
