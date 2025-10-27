#!/usr/bin/env python3
"""Run extractor with LLM formatting"""

import subprocess
import sys

print("Running PowerPoint extractor with LLM formatting...")
print("=" * 60)

result = subprocess.run(
    [sys.executable, "ppt_extractor.py", "MCIT AI_Unified Platform.pptx"],
    capture_output=True,
    text=True,
    cwd="."
)

print("\nSTDOUT:")
print(result.stdout)

print("\nSTDERR:")
print(result.stderr)

print(f"\nReturn code: {result.returncode}")

