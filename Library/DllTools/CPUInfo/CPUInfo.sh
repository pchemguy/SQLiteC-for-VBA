#!/bin/bash
#
# MSYS/MinGW build script
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

readonly __SCRIPT_NAME_FULL__="${0##*/}"
readonly __SCRIPT_NAME_BASE__="${__SCRIPT_NAME_FULL__%.*}"


main() {
  if [[ "${MSYSTEM}" == "MINGW64" ]]; then
    readonly ARCH="x64"
  else
    readonly ARCH="x32"
  fi
  
  case "${1:-0}" in
    "0")
      OptLevel=0
      ;;
    "1")
      OptLevel=1
      ;;
    "2")
      OptLevel=2
      ;;
    "3")
      OptLevel=3
      ;;
    *)
      OptLevel=0
      ;;
  esac

  mkdir -p "./${ARCH}"

  readonly SrcName="${__SCRIPT_NAME_BASE__}"

  gcc -O0 -Wall -msse ${SrcName}.c -o ${SrcName}.exe

  mv ${SrcName}.exe "./${ARCH}"

  return 0
}


main "$@"
exit 0
