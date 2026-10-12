#!/usr/bin/env bash
set -euo pipefail
xmlmcp_script_dir="$(cd -P -- "$(dirname -- "$0")" >/dev/null && pwd)"
xmlmcp_python="${XMLMCP_PYTHON:-python3}"
exec "$xmlmcp_python" "$xmlmcp_script_dir/setup.py" "$@"
