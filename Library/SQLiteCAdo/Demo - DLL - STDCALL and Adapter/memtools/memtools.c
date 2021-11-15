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
}
