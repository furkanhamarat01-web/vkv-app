#!/bin/bash
cd "$(dirname "$0")" || exit 1
for PY in /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3; do
  if [ -x "$PY" ]; then
    exec "$PY" "$PWD/istemci.py"
  fi
done
echo "Python 3 bulunamadı."
exit 1
