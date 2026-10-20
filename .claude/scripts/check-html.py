#!/usr/bin/env python3
"""Extract inline <script> blocks from HTML files and run `node --check` on each.

Usage: python3 check_html.py file1.html [file2.html ...]
Exit code 1 if any block fails.
"""
import re
import subprocess
import sys
import tempfile
import os

SCRIPT_RE = re.compile(r"<script(?P<attrs>[^>]*)>(?P<body>.*?)</script>", re.S | re.I)

failed = False
for path in sys.argv[1:]:
    html = open(path, encoding="utf-8").read()
    blocks = [m for m in SCRIPT_RE.finditer(html) if "src=" not in m.group("attrs")]
    for i, m in enumerate(blocks, 1):
        body = m.group("body")
        if not body.strip():
            continue
        is_module = 'type="module"' in m.group("attrs") or "type='module'" in m.group("attrs")
        suffix = ".mjs" if is_module else ".js"
        line = html[: m.start("body")].count("\n") + 1
        with tempfile.NamedTemporaryFile("w", suffix=suffix, delete=False, encoding="utf-8") as f:
            f.write(body)
            tmp = f.name
        r = subprocess.run(["node", "--check", tmp], capture_output=True, text=True)
        os.unlink(tmp)
        status = "OK" if r.returncode == 0 else "FAIL"
        print(f"{status}  {path}  script#{i} ({'module' if is_module else 'classic'}, starts line {line})")
        if r.returncode != 0:
            failed = True
            print(r.stderr.strip())
sys.exit(1 if failed else 0)
