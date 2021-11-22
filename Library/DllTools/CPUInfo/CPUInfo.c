#include <stdio.h>
#include <time.h>
#include <intrin.h>


typedef struct CPUGeneralRegisters {
  unsigned int eax;
  unsigned int ebx;
  unsigned int ecx;
  unsigned int edx;
} CPUGeneralRegisters;


void GetCPUFreq(CPUGeneralRegisters *CPUGenRegs);
static int inline ReferenceCPULoad(int CycleCount);


void GetCPUFreq(CPUGeneralRegisters *CPUGenRegs) {
  __cpuid((int *) CPUGenRegs, 0);
  if (CPUGenRegs->eax >= 0x16) {
    __cpuid((int *) CPUGenRegs, 0x16);
  } else {
    CPUGenRegs->eax = 0;
  }
}


static int inline ReferenceCPULoad(int CycleCount) {
    __m128 x = _mm_setzero_ps();
    for(int i=0; i < CycleCount; i++) {
        x = _mm_add_ps(x, _mm_set1_ps(1.0f));
    }
    return _mm_cvt_ss2si(x);
}


int main(void) {
  CPUGeneralRegisters CPUGenRegs = (CPUGeneralRegisters) {0, 0, 0, 0};
  GetCPUFreq(&CPUGenRegs);

  if (CPUGenRegs.eax > 0) {
    printf("EAX: 0x%08x EBX: 0x%08x ECX: %08x\r\n",
           CPUGenRegs.eax, CPUGenRegs.ebx, CPUGenRegs.ecx);
    printf("Processor Base Frequency:  %04d MHz\r\n", CPUGenRegs.eax);
    printf("Maximum Frequency:         %04d MHz\r\n", CPUGenRegs.ebx);
    printf("Bus (Reference) Frequency: %04d MHz\r\n", CPUGenRegs.ecx);
  } else {
    printf("CPUID level 16h unsupported\r\n");
  }

  struct timeb start, end;

  ftime(&start);
  ReferenceCPULoad(1e9);
  ftime(&end);

  const int MSEC_IN_SEC = 1000;
  int diff;
  diff = MSEC_IN_SEC * (end.time - start.time) + (end.millitm - start.millitm);

  printf("\nReferenceCPULoad - 1e9 times - %u milliseconds\n", diff);
  
  
  return 0;
}
