#!/bin/bash
echo "Starting Excel Analyzer MCP Server..."
cd "$(dirname "$0")"
export PATH="$(pwd)/.venv/bin:$PATH"
.venv/bin/python src/excel_analyzer.py
