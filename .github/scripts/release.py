"""Coordinated version metadata and fail-closed release publication."""

import argparse
import os
from pathlib import Path
import plistlib
import re
import subprocess
import sys
from urllib.error import HTTPError
from urllib.request import Request, urlopen


PROJECT = Path("msoffice crypt.xcodeproj/project.pbxproj")
PLIST = Path("Info.plist")
LOCALIZED = Path("English.lproj/InfoPlist.strings")
SEMVER = re.compile(r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)", re.ASCII)
SETTINGS = ("MARKETING_VERSION", "CURRENT_PROJECT_VERSION")
PLIST_SETTINGS = {
    "CFBundleShortVersionString": "MARKETING_VERSION",
    "CFBundleVersion": "CURRENT_PROJECT_VERSION",
}


def require_version(value):
    if not isinstance(value, str) or not SEMVER.fullmatch(value):
        raise ValueError(f"Invalid semantic version: {value!r}; expected X.Y.Z without leading zeros")
    return value


def assignments(text, key):
    return re.findall(rf"\b{key}\s*=\s*([^;]+);", text)


def setting_value(value):
    value = value.strip()
    if re.fullmatch(r'"[^"]*"', value):
        return value[1:-1]
    return value


def validate_plist(data, version, allow_variables=True):
    for key, setting in PLIST_SETTINGS.items():
        if key not in data:
            if key == "CFBundleVersion":
                continue  # This project's source plist has no build-version key.
            raise ValueError(f"Missing {key}")
        value = data[key]
        if allow_variables and value in (f"$({setting})", f"${{{setting}}}"):
            continue
        if require_version(value) != version:
            raise ValueError(f"{key} drift: {value!r} != {version}")


def read_version():
    project = PROJECT.read_text()
    versions = []
    blocks = re.findall(r"\bbuildSettings = \{(.*?)\};", project, re.S)
    if not blocks:
        raise ValueError("No Xcode build settings found")
    for block in blocks:
        for key in SETTINGS:
            values = assignments(block, key)
            if len(values) != 1:
                raise ValueError(f"Expected one {key} in every Xcode build configuration")
            versions.append(require_version(setting_value(values[0])))
    if len(set(versions)) != 1:
        raise ValueError(f"Xcode version drift: {versions}")
    version = versions[0]
    with PLIST.open("rb") as source:
        validate_plist(plistlib.load(source), version)
    localized = LOCALIZED.read_text()
    for key in ("CFBundleShortVersionString", "CFBundleGetInfoString"):
        values = assignments(localized, key)
        if len(values) != 1:
            raise ValueError(f"Expected one localized {key}")
        value = setting_value(values[0])
        if key == "CFBundleGetInfoString":
            match = re.fullmatch(r"msoffice crypt version ([^,]+), Copyright 2021 miyako\.", value)
            if not match:
                raise ValueError(f"Unrecognized localized {key}: {value!r}")
            value = match[1]
        if require_version(value) != version:
            raise ValueError(f"Localized {key} drift: {value} != {version}")
    return version


def bump(mode):
    current = read_version()
    parts = list(map(int, current.split(".")))
    index = ("major", "minor", "patch").index(mode)
    parts[index] += 1
    parts[index + 1:] = [0] * (2 - index)
    new = ".".join(map(str, parts))
    project = PROJECT.read_text()
    for key in SETTINGS:
        project = re.sub(rf"(\b{key}\s*=\s*)[^;]+;", rf"\g<1>{new};", project)
    PROJECT.write_text(project)
    # Preserve build-variable references and the plist's existing formatting.
    plist = PLIST.read_text()
    for key in PLIST_SETTINGS:
        plist = re.sub(
            rf"(<key>{key}</key>\s*<string>){re.escape(current)}(</string>)",
            rf"\g<1>{new}\2", plist,
        )
    PLIST.write_text(plist)
    localized = LOCALIZED.read_text()
    localized = localized.replace(f'"{current}"', f'"{new}"')
    localized = localized.replace(f"version {current},", f"version {new},")
    LOCALIZED.write_text(localized)
    if read_version() != new:
        raise ValueError("Bumped metadata did not validate")
    return new


def ensure_available(tag):
    result = subprocess.run(
        ["git", "ls-remote", "--exit-code", "--tags", "origin", f"refs/tags/{tag}"],
        text=True, capture_output=True, check=False,
    )
    if result.returncode == 0:
        raise ValueError(f"Tag {tag} already exists; bump the version before building")
    if result.returncode != 2:
        raise RuntimeError(f"Cannot check remote tags: {result.stderr}")
    request = Request(
        f"{os.environ['GITHUB_API_URL']}/repos/{os.environ['GITHUB_REPOSITORY']}/releases/tags/{tag}",
        headers={
            "Authorization": f"Bearer {os.environ['GH_TOKEN']}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urlopen(request, timeout=30):
            pass
    except HTTPError as error:
        if error.code == 404:
            return
        raise
    raise ValueError(f"Release {tag} already exists; refusing to replace it or its assets")


def check_commit():
    sha = os.environ["GITHUB_SHA"]
    if not re.fullmatch(r"[0-9a-f]{40}", sha):
        raise ValueError(f"Invalid selected commit: {sha!r}")
    actual = subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip()
    if actual != sha:
        raise ValueError(f"Checkout {actual} does not match selected commit {sha}")
    return sha


def publish(version, assets):
    sha = check_commit()
    if read_version() != require_version(version):
        raise ValueError("Publication version differs from selected commit metadata")
    for asset in assets:
        if not Path(asset).is_file() or Path(asset).stat().st_size == 0:
            raise ValueError(f"Missing or empty release asset: {asset}")
    tag = f"v{version}"
    ensure_available(tag)
    subprocess.run(["git", "tag", tag, sha], check=True)
    subprocess.run([
        "git", "-c", "push.followTags=false", "push", "origin",
        f"refs/tags/{tag}:refs/tags/{tag}",
    ], check=True)
    remote = subprocess.check_output(
        ["git", "ls-remote", "--exit-code", "--tags", "origin", f"refs/tags/{tag}"], text=True,
    ).split()
    if remote != [sha, f"refs/tags/{tag}"]:
        raise ValueError("Published tag does not point to the selected commit")
    subprocess.run([
        "gh", "release", "create", tag, "--verify-tag", "--target", sha,
        "--title", tag, "--generate-notes", *assets,
    ], check=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("version")
    commands.add_parser("preflight")
    bump_parser = commands.add_parser("bump")
    bump_parser.add_argument("mode", choices=("patch", "minor", "major"))
    verify = commands.add_parser("verify-bundle")
    verify.add_argument("plist", type=Path)
    release = commands.add_parser("publish")
    release.add_argument("version")
    release.add_argument("assets", nargs=2)
    args = parser.parse_args()
    if args.command == "bump":
        version = bump(args.mode)
    elif args.command == "publish":
        publish(args.version, args.assets)
        return
    else:
        version = read_version()
        if args.command == "preflight":
            check_commit()
            ensure_available(f"v{version}")
        elif args.command == "verify-bundle":
            with args.plist.open("rb") as source:
                validate_plist(plistlib.load(source), version, allow_variables=False)
    print(version)
    if output := os.environ.get("GITHUB_OUTPUT"):
        with open(output, "a") as destination:
            destination.write(f"version={version}\ntag=v{version}\n")


if __name__ == "__main__":
    try:
        main()
    except (ValueError, RuntimeError, OSError, subprocess.CalledProcessError) as error:
        print(f"::error::{error}", file=sys.stderr)
        sys.exit(1)
