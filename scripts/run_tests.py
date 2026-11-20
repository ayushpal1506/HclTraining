"""Run the suite once and write the evidence file the documents quote.

The Step 13 document, the Step 15 report and the viva all need the same answer
to "what actually passed?" — so the answer is produced by one script and stored
as JSON, not retyped into a document.

    .venv/bin/python scripts/run_tests.py

Writes data/test_results.json and prints a one-screen summary.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
XML = ROOT / "data" / "pytest_junit.xml"
OUT = ROOT / "data" / "test_results.json"
ID_RE = re.compile(r"TC[_-]T[_-](\d{2})([a-z]?)")


def main() -> int:
    cmd = [sys.executable, "-m", "pytest", "tests", "-q",
           "--junitxml", str(XML), "--durations=5"]
    proc = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True)
    print(proc.stdout[-4000:])
    if proc.stderr.strip():
        print(proc.stderr[-2000:], file=sys.stderr)

    if not XML.exists():
        print("no JUnit XML produced — pytest could not start", file=sys.stderr)
        return 1

    suite = ET.parse(XML).getroot().find("testsuite")
    counts = {k: int(suite.get(k, 0)) for k in
              ("tests", "failures", "errors", "skipped")}
    seconds = float(suite.get("time", 0.0))

    cases: dict[str, str] = {}
    names: dict[str, str] = {}
    slowest = []
    for case in suite.iter("testcase"):
        name = case.get("name", "")
        seconds_case = float(case.get("time", 0.0))
        slowest.append((round(seconds_case, 2), name))
        outcome = "PASS"
        if case.find("failure") is not None:
            outcome = "FAIL"
        elif case.find("error") is not None:
            outcome = "ERROR"
        elif case.find("skipped") is not None:
            outcome = "SKIP"
        match = ID_RE.search(name)
        if match:
            key = f"TC-T-{match.group(1)}{match.group(2)}"
            cases[key] = outcome
            names[key] = name

    # A test can be parametrised into several cases; an outright failure wins.
    for case in suite.iter("testcase"):
        match = ID_RE.search(case.get("name", ""))
        if not match:
            continue
        key = f"TC-T-{match.group(1)}{match.group(2)}"
        if cases.get(key) == "PASS" and any(
                case.find(tag) is not None for tag in ("failure", "error", "skipped")):
            cases[key] = ("FAIL" if case.find("failure") is not None else
                          "ERROR" if case.find("error") is not None else "SKIP")

    payload = {
        "generated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "python": sys.version.split()[0],
        "command": " ".join(cmd[1:]),
        "counts": counts,
        "seconds": round(seconds, 1),
        "cases": dict(sorted(cases.items())),
        "titles": names,
        "slowest": sorted(slowest, reverse=True)[:5],
    }
    OUT.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    tally = Counter(cases.values())
    print(f"\n{counts['tests']} test items · {counts['failures']} fail · "
          f"{counts['errors']} error · {counts['skipped']} skip · {seconds:.1f} s")
    print(f"TC-T ids recorded: {len(cases)}  {dict(sorted(tally.items()))}")
    print(f"evidence: {OUT.relative_to(ROOT)}")
    return 0 if counts["failures"] == 0 and counts["errors"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
