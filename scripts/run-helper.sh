#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")/../helper"
exec .venv/bin/python -m gurobistudio
