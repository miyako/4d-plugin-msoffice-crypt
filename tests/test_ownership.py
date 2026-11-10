#!/usr/bin/env python3
"""Compile actual helpers and the unchanged SDK; never build a plugin bundle."""

import argparse
import os
from pathlib import Path
import platform
import subprocess
import tempfile


ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", help="Immutable git commit expected to fail setters and getters")
    args = parser.parse_args()
    if platform.system() != "Darwin":
        parser.error("This harness uses the real macOS CoreFoundation conversion path")
    with tempfile.TemporaryDirectory(prefix="office-crypt-ownership-") as directory:
        work = Path(directory)
        helper = ROOT / "support/4DPlugin-JSON.cpp"
        if args.baseline:
            sha = subprocess.check_output(
                ["git", "rev-parse", "--verify", args.baseline + "^{commit}"],
                cwd=ROOT, text=True,
            ).strip()
            helper = work / "4DPlugin-JSON.cpp"
            helper.write_bytes(subprocess.check_output(
                ["git", "show", sha + ":support/4DPlugin-JSON.cpp"], cwd=ROOT,
            ))
            print("Baseline:", sha, flush=True)
        binary = work / "ownership"
        subprocess.run([
            "clang++", "-std=c++11", "-g", "-O1",
            "-fsanitize=address,undefined", "-fno-omit-frame-pointer",
            "-Wno-deprecated-declarations", "-Wno-multichar",
            "-I" + str(ROOT / "4D Plugin API"), "-I" + str(ROOT / "support"),
            "-x", "c++", str(ROOT / "4D Plugin API/4DPluginAPI.c"),
            str(helper), str(ROOT / "tests/ownership.cpp"),
            "-framework", "CoreFoundation", "-framework", "CoreGraphics",
            "-o", str(binary),
        ], check=True)
        environment = dict(os.environ, ASAN_OPTIONS="detect_leaks=0", UBSAN_OPTIONS="halt_on_error=1")
        for case in ("setters", "getters", "arrays"):
            result = subprocess.run([str(binary), case], env=environment,
                                    capture_output=True, text=True)
            expected_failure = bool(args.baseline and case != "arrays")
            if expected_failure:
                if result.returncode == 0 or "ownership imbalance" not in result.stderr:
                    raise SystemExit(f"{case}: expected ownership-accounting failure\n"
                                     + result.stdout + result.stderr)
                print(f"{case}: expected baseline failure\n{result.stderr.strip()}", flush=True)
            else:
                print(result.stdout, end="", flush=True)
                if result.returncode:
                    raise SystemExit(result.stderr)


if __name__ == "__main__":
    main()
