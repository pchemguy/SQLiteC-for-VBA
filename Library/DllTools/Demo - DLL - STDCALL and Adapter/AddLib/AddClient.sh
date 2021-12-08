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

  [[ ! -r "./${ARCH}/addlib.dll" ]] && echo "addlib.dll not found." && exit 101
  
  # Only use -DADD_EXPORTS when compiling the library
  gcc -c addclient.c -o addclient.o
  gcc addclient.o -o addclient.exe -L"./${ARCH}" -laddlib

  rm addclient.o
  mv addclient.exe "./${ARCH}"

  return 0
}


main "$@"
exit 0
