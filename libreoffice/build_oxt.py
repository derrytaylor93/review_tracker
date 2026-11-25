#!/usr/bin/env python3
"""Package the extension folder into ReviewTracker.oxt (a zip file)."""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "extension")
OUT = os.path.join(HERE, "ReviewTracker.oxt")

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(SRC):
        for f in sorted(files):
            if f.endswith((".pyc",)) or f == ".DS_Store":
                continue
            full = os.path.join(root, f)
            z.write(full, os.path.relpath(full, SRC))
print("Built", OUT)
