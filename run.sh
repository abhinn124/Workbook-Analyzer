#!/bin/bash
echo "Starting Excel Analyzer MCP Server..."
cd "$(dirname "$0")"
.venv/bin/python src/excel_analyzer.py