#include "memtools.h"


MEMTOOLSAPI void MEMTOOLSCALL CopyMem(void* Destination, const void* Source, size_t Length) {
  switch(Length) {
    case 4:
      *(int32_t*)Destination = *(int32_t*)Source;
      break;
    case 8:
      *(int64_t*)Destination = *(int64_t*)Source;
      break;
    case 0:
      break;
    case 1:
      *(int8_t*)Destination = *(int8_t*)Source;
      break;
    case 2:
      *(int16_t*)Destination = *(int16_t*)Source;
      break;
    default:
      memcpy(Destination, Source, Length);
      break;
  }
  return;
}

// Volatile loop counter should be used here to prevent optimization.
MEMTOOLSAPI int MEMTOOLSCALL PerfGauge(unsigned int ForCount) {
  struct timeb start, end;
  ftime(&start);
  for (volatile unsigned int i=0; i < ForCount; i++) {
    ;
  }
  ftime(&end);
  return 1000 * (end.time - start.time) + (end.millitm - start.millitm);
}

MEMTOOLSAPI void MEMTOOLSCALL DummySub0Args() {
  return;
}

MEMTOOLSAPI void MEMTOOLSCALL DummySub3Args(void* Destination, const void* Source, size_t Length) {
  return;
}

MEMTOOLSAPI int MEMTOOLSCALL DummyFnc0Args() {
  volatile int Result = 10241024;
  return Result;
}

MEMTOOLSAPI int MEMTOOLSCALL DummyFnc3Args(void* Destination, const void* Source, size_t Length) {
  volatile int Result = 10241024;
  return Result;
}
