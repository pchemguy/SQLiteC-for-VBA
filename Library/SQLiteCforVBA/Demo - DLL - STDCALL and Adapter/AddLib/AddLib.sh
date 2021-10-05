#!/bin/sh
#
set -euo pipefail
IFS=$'\n\t'

# Build STDCALL version:
# $ ABI=STDCALL ./sqlite3.ref.sh

cleanup_EXIT() { 
  echo "EXIT clean up: $?" 
}
trap cleanup_EXIT EXIT

cleanup_TERM() {
  echo "TERM clean up: $?"
}
trap cleanup_TERM TERM

cleanup_ERR() {
  echo "ERR clean up: $?"
}
trap cleanup_ERR ERR

#gcc -shared -o sqlite3.dll sqlite3.c -Wl,--output-def,sqlite3.def,--out-implib,libsqlite3.a

#gcc -o sqlite3.dll sqlite3.c -shared -s -Wl,--subsystem,windows,--output-def,sqlite3.def
#gcc -o sqlite3.dll sqlite3.c -shared -s -Wl,--subsystem,windows,--kill-at
#dlltool --kill-at -d sqlite3.def -D qlite3.dll --output-lib libsqlite3.a

gcc -mrtd -O3 -std=c99 -Wall -c add.c -o add.o
gcc -o AddLib.dll add.o -shared -s -Wl,--subsystem,windows,--output-def,AddLib.def
gcc -o AddLib.dll add.o -shared -s -Wl,--subsystem,windows,--kill-at
dlltool --kill-at -d AddLib.def -D AddLib.dll -l libaddlib.a

