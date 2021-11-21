---
layout: default
title: VBA to DLL call efficiency
nav_order: 3
parent: DLL Manager
permalink: /dllman/vba-dll-call
---

While playing with a [VBA project][VBA-MemoryTools]  involving Windows APIs, I have noticed some odd behavior: API calls under Excel 2016/VBA7/x64 appeared to be much slower than under Excel 2002/VBA6/x32. Since I wanted to follow the recipe described in the [Weak Reference][] post, I decided to look into this performance matter more closely.

### DLL side: Mock DLL library with test fixtures and a C-client

To simplify the interpretation of results, I created a small C-based dll with several stubs (memtools.c) and a C-based client calling this dll and timing the calls (memtoolsclient.c). (The sources, MSYS/MinGW build scripts, and precompiled binaries are in the *Library/DllTools/Memtools* folder.) This C-code to C-code call timings provide performance references. The DLL includes five stubs. The first function, PerfGauge, takes a loop counter, times execution of an empty For loop within the DLL, and returns the result. The remaining stubs permit timing the dll calls and examine differences associated with passing arguments and returning a value. As their names suggest, these stubs

* are either a Sub or Function and
* take either zero or three arguments.

### VBA side: DllPerfLib class with VBA fixtures

The DllPerfLib VBA class calls these stubs and times them. Out of curiosity, I also added VBA stubs with signatures matching those in the dll so that DllPerfLib also times calls to these stubs yielding performance of native calls. DllPerfLib includes several blocks. The first block starting from the top contains dll stub declarations, factory, constructor, and attribute accessors. The second block handles dll loading from the project directory via the [DllManager][] class. Then follows a wrapper for the performance reference routine (empty timed For loop inside the dll). The next block includes four functions, timing calls to corresponding dll stubs and their VBA twins defined in the last section of the class.

### Test results

I compiled x32 and x64 versions of the DLL and the C-client with MSYS/MinGW toolchains on Windows with no optimization (-O0), ran tests on each pair, and used the two DLL's to run tests from Excel 2002/VBA6/x32 and Excel 2016/VBA7/x64 respectively. I ran DllPerfRun.Runner multiple times, discarded results that were much slower than the rest, and calculated average timings. [Table 1](#VBADLLPerformance) shows a representative subset of results.

<div align="center"><b>Table 1. Time in seconds required for completion of 10<sup>9</sup> repetitions.</b></div>
<a name="VBADLLPerformance"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/SQLiteC-for-VBA/develop/Assets/Diagrams/VBA%20Performance.svg" alt="VBADLLPerformance" width="75%" /></div>

#### C-client timings

The leftmost column of [Table 1](#VBADLLPerformance) contains the PerfGauge timing. While I am not examining the disassembled code (which is a prudent thing to do), an empty unoptimized C-language For loop should require at least three machine instructions:

* increment the loop variable,
* compare the loop variable with the target,
* perform a conditional jump.

On a 2.2 GHz multi-core processor with dynamic frequency adjustment (Intel Core i7-8750H @2.2GHz), the number of 2.1 s for 10<sup>9</sup> repetitions, therefore, appears to be qualitatively reasonable. The second column shows the timing for calling DummySub0Args from the C-client (see DummySub0ArgsGauge routine). Explaining why the numbers in the second column are lower would require further investigation. However, I am more concerned about the results in the right half of the table.

#### VBA timings

The green cell highlights the efficiency of calling a DLL routine from VBA6/x32/Excel 2002. This number indicates that a DLL call taking no arguments and returning no value is only 5x times slower than the same call from a compiled C-client. Further, this call is 7x times faster than a native VBA call (rightmost column) with the same signature. When the called routine either takes arguments or returns a value, the difference is less pronounced. Still, with the other three implemented mock calls, the tendency is qualitatively similar.

The primary concern is the cell with an orange background, indicating that a single DLL call takes 2 microseconds under 2016/VBA7/x64 instead of 8 nanoseconds under VBA6/x32/Excel 2002.


<!-- References -->

[VBA-MemoryTools]: https://codereview.stackexchange.com/questions/252659/fast-native-memory-manipulation-in-vba
[Weak Reference]: https://rubberduckvba.wordpress.com/2018/09/11/lazy-object-weak-reference/
[DllManager]: https://codereview.stackexchange.com/questions/268630/
