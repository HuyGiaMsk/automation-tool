#!/bin/bash

if [ ! -d venv ]; then
  python -m venv venv
  echo "Virtual environment created: venv"
fi

source venv/bin/activate

pip install -r requirements.txt

pyinstaller automation_tool.spec

if [ -d dist ]; then
  cp -r input script dist/
fi

deactivate
