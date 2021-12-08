#!/bin/bash
#
set -euo pipefail
IFS=$'\n\t'

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


main() {
  if [[ "${MSYSTEM}" == "MINGW64" ]]; then
    readonly ARCH="x64"
  else
    readonly ARCH="x32"
  fi

  rm -rf "./${ARCH}"
  mkdir -p "./${ARCH}"
  
  # Only use -DADD_EXPORTS when compiling the library
  gcc -O3 -Wall -c add.c -o add.o -DADD_EXPORTS
  gcc -o AddLib.dll add.o -shared -Wl,--subsystem,windows,--output-def,AddLib.def
  gcc -o AddLib.dll add.o -shared -Wl,--subsystem,windows,--kill-at
  dlltool --kill-at -d AddLib.def -D AddLib.dll -l libaddlib.a

  rm add.o
  mv AddLib.d* "./${ARCH}"
  mv libaddlib.a "./${ARCH}"

  return 0
}


main "$@"
exit 0
