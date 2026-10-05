#!/usr/bin/env bash
# Convenience shim: run bli2prism without setting PYTHONPATH.
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHONPATH="$here${PYTHONPATH:+:$PYTHONPATH}" exec python3 -m bli2prism "$@"
