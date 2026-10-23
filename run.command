#!/bin/bash
# Mac/Linux: double-click or run ./run.command
cd "$(dirname "$0")"
[ -d .venv ] || { python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt; }
(sleep 2; open http://localhost:5055 2>/dev/null || xdg-open http://localhost:5055) &
.venv/bin/python app.py
