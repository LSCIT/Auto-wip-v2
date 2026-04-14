#!/usr/bin/env python3
"""Split olevba output into individual module files."""
import re
import os

INPUT = "vba_source/full_vba_dump.txt"
OUTPUT_DIR = "vba_source"

with open(INPUT, "r", encoding="utf-8", errors="replace") as f:
    content = f.read()

# Pattern: VBA MACRO <name> \n in file: ... \n - - - - \n <code>
# Split on the VBA MACRO header lines
pattern = r"VBA MACRO (\S+)\s*\nin file:.*?\n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - \n"
parts = re.split(pattern, content)

# parts[0] is header, then alternating: name, code, name, code...
modules = {}
for i in range(1, len(parts) - 1, 2):
    name = parts[i]
    code = parts[i + 1]
    # Trim trailing analysis sections
    code = re.split(r"\n-{10,}\n", code)[0]
    modules[name] = code

for name, code in modules.items():
    filepath = os.path.join(OUTPUT_DIR, name)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(code)
    lines = code.count('\n')
    print(f"  {name}: {lines} lines")

print(f"\nExtracted {len(modules)} modules")
