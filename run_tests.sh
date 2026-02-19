#!/bin/bash
cd "$(dirname "$0")"

if [ -d "venv14" ]; then
    source venv14/bin/activate
fi

pytest testing/test_pmojo.py -v "$@"
