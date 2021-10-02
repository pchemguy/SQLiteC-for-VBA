rem Build SQLite using MSVC toolset
@echo off

if /%Platform%/==/x86/ (
  set ARCH=
) else (
  set ARCH=64
)

set USE_STDCALL=1
set SESSION=1
set RBU=1
set NO_TCL=1
set USE_WINV100_NSDKLIBPATH=1

rem For now ICU is disabled.
rem Compilation against precompiled MSVC 2019 binaries completes OK, but
rem the resulting library could not be loaded (ICU dll's are placed in
rem the same folder as the library). Attempt to compile ICU from source
rem failed. Further investigation of these issues is necessary.
set USE_ICU=0
rem set ICUDIR=%ProgramFiles%\icu4c
rem set ICUDIR=%ICUDIR: =%
rem set ICUINCDIR=%ICUDIR%\include
rem set ICULIBDIR=%ICUDIR%\lib%ARCH%

REM set USE_ZLIB=1
REM set ZLIBDIR=..\zlib


set EXT_FEATURE_FLAGS=^
-DSQLITE_ENABLE_FTS3_PARENTHESIS ^
-DSQLITE_ENABLE_FTS3_TOKENIZER ^
-DSQLITE_ENABLE_FTS4=1 ^
-DSQLITE_ENABLE_FTS5=1 ^
-DSQLITE_SYSTEM_MALLOC=1 ^
-DSQLITE_OMIT_LOCALTIME=1 ^
-DSQLITE_DQS=0 ^
-DSQLITE_LIKE_DOESNT_MATCH_BLOBS ^
-DSQLITE_MAX_EXPR_DEPTH=100 ^
-DSQLITE_OMIT_DEPRECATED ^
-DSQLITE_DEFAULT_FOREIGN_KEYS=1 ^
-DSQLITE_DEFAULT_SYNCHRONOUS=1 ^
-DSQLITE_ENABLE_EXPLAIN_COMMENTS ^
-DSQLITE_ENABLE_OFFSET_SQL_FUNC=1 ^
-DSQLITE_ENABLE_QPSG ^
-DSQLITE_ENABLE_STMTVTAB ^
-DSQLITE_ENABLE_STAT4 ^
-DSQLITE_SOUNDEX

if not /%~1/==// (
  set TARGET=%~1
) else (
  set TARGET=echoconfig
)


rem Generates "splitline.bat" script.
rem "splitline.bat" takes one quoted argument, splits it on the
rem space character, and outputs each part on a separate line.
echo ========== Generating "splitline.bat" ==========
set OUTPUT="splitline.bat"
1>%OUTPUT% (
  echo @echo off
)
1>>%OUTPUT% (
  echo.
  echo set ARGS=%%~1
  echo :NEXT_ARG
  echo   for /F "tokens=1* delims= " %%%%G in ^("%%ARGS%%"^) do ^(
  echo     echo %%%%G
  echo     set ARGS=%%%%H
  echo   ^)
  echo if defined ARGS goto NEXT_ARG
)


cd /d %~dp0sqlite

if not exist Makefile.msc.bak (
  copy Makefile.msc Makefile.msc.bak
) else (
  copy /Y Makefile.msc.bak Makefile.msc 1>nul
)

echo ========== Patching   "Makefile.msc" ===========
set OUTPUT="Makefile.msc"
set "TAB=	"
1>>%OUTPUT% (
  echo.
  echo echoconfig:
  echo %TAB%@echo --------------------------------
  echo %TAB%@echo REQ_FEATURE_FLAGS
  echo %TAB%@splitline.bat "$(REQ_FEATURE_FLAGS)"
  echo %TAB%@echo --------------------------------
  echo %TAB%@echo OPT_FEATURE_FLAGS
  echo %TAB%@splitline.bat "$(OPT_FEATURE_FLAGS)"
  echo %TAB%@echo --------------------------------
  echo %TAB%@echo EXT_FEATURE_FLAGS
  echo %TAB%@splitline.bat "$(EXT_FEATURE_FLAGS)"
  echo %TAB%@echo --------------------------------
  echo %TAB%@echo TCC
  echo %TAB%@splitline.bat "$(TCC)"
  echo %TAB%@echo --------------------------------
  echo %TAB%@echo USE_STDCALL=$^(USE_STDCALL^)
  echo %TAB%@echo USE_ZLIB=$^(USE_ZLIB^)
  echo %TAB%@echo USE_ICU=$^(USE_ICU^)
  echo %TAB%@echo FOR_WIN10=$^(FOR_WIN10^)
  echo %TAB%@echo DEBUG=$^(DEBUG^)
  echo %TAB%@echo SESSION=$^(SESSION^)
  echo %TAB%@echo RBU=$^(RBU^)
  echo %TAB%@echo ICUDIR=$^(ICUDIR^)
  echo %TAB%@echo ICUINCDIR=$^(ICUINCDIR^)
  echo %TAB%@echo ICULIBDIR=$^(ICULIBDIR^)
  echo %TAB%@echo ZLIBDIR=$^(ZLIBDIR^)
)


nmake /nologo /f Makefile.msc %TARGET%
cd ..


