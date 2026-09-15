#!/bin/zsh
set -e
cd "${0:A:h}"
PYTHON_BIN="${PDD_PYTHON:-python3}"
"$PYTHON_BIN" -m unittest discover -s tests -v
