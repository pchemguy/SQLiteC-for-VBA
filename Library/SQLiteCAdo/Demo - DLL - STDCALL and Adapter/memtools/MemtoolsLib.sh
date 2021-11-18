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

  mkdir -p "./${ARCH}"

  readonly SrcName="memtools"
  rm -f "./${ARCH}/${SrcName}lib.d"*
  rm -f "./${ARCH}/lib${SrcName}lib.a"
  
  # Only use -Dxxx_EXPORTS when compiling the library
  gcc -O3 -Wall -c ${SrcName}.c -o ${SrcName}.o -DMEMTOOLS_EXPORTS
  gcc -o ${SrcName}lib.dll ${SrcName}.o -shared -Wl,--subsystem,windows,--output-def,${SrcName}lib.def
  gcc -o ${SrcName}lib.dll ${SrcName}.o -shared -Wl,--subsystem,windows,--kill-at
  dlltool --kill-at -d ${SrcName}lib.def -D ${SrcName}lib.dll -l lib${SrcName}lib.a

  rm -f ${SrcName}.o
  mv -f ${SrcName}lib.d* "./${ARCH}"
  mv -f lib${SrcName}lib.a "./${ARCH}"

  return 0
}


main "$@"
exit 0
