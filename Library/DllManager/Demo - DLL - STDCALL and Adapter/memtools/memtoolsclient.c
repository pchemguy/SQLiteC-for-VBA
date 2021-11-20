#include "memtools.h"


void CopyMemGauge();
void DummySub0ArgsGauge();
void DummySub3ArgsGauge();
void DummyFnc0ArgsGauge();
void DummyFnc3ArgsGauge();


int main(int argc, char** argv) { 
  CopyMemGauge();
  DummySub0ArgsGauge();
  DummySub3ArgsGauge();
  DummyFnc0ArgsGauge();
  DummyFnc3ArgsGauge();
  return 0;
}


void CopyMemGauge() {
  struct timeb start, end;

  int dest;
 
  ftime(&start);
  for (int i=0; i < 1e8; i++) {
    CopyMem(&dest, &i, sizeof dest); 
  }
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nCopyMemGauge - 1e8 times - %u milliseconds\n", diff);
  printf("\nFinal value: %u\n", dest);
}


void DummySub0ArgsGauge() {
  void (*volatile MEMTOOLSCALL pDummySub0Args)();
  pDummySub0Args = DummySub0Args;

  struct timeb start, end;

  ftime(&start);
  for (volatile int i=0; i < 1e9; i++) {
    pDummySub0Args(); 
  }
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nDummySub0Args - 1e9 times - %u milliseconds\n", diff);
}


void DummySub3ArgsGauge() {
  char Src[] = "ABCDEFGHIJKLMNOPGRSTUVWXYZABCDEFGHIJKLMNOPGRSTUVWXYZ";
  char Dst[255];
  size_t SrcLen = sizeof(Src);

  struct timeb start, end;

  ftime(&start);
  for (volatile int i=0; i < 1e9; i++) {
    DummySub3Args(Dst, Src, SrcLen);
  }
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nDummySub3Args - 1e9 times - %u milliseconds\n", diff);
}


void DummyFnc0ArgsGauge() {
  int Result __attribute__((unused));
  struct timeb start, end;

  ftime(&start);
  for (volatile int i=0; i < 1e9; i++) {
    Result = DummyFnc0Args(); 
  }
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nDummyFnc0Args - 1e9 times - %u milliseconds\n", diff);
}


void DummyFnc3ArgsGauge() {
  char Src[] = "ABCDEFGHIJKLMNOPGRSTUVWXYZABCDEFGHIJKLMNOPGRSTUVWXYZ";
  char Dst[255];
  size_t SrcLen = sizeof(Src);

  int Result __attribute__((unused));
  struct timeb start, end;

  ftime(&start);
  for (volatile int i=0; i < 1e9; i++) {
    Result = DummyFnc3Args(Dst, Src, SrcLen);
  }
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nDummyFnc3Args - 1e9 times - %u milliseconds\n", diff);
}
