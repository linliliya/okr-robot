"""Assemble the existing site and the separately maintained AI schedule."""

import json
import os
from pathlib import Path
import shutil
import subprocess
import time
import urllib.error
import urllib.request


ROOT = Path(__file__).resolve().parents[2]
SITE = ROOT / "_site"


def git(*args, cwd=ROOT):
    return subprocess.check_output(["git", *args], cwd=cwd).decode().strip()


def assemble():
    versions = {
        "portal": git("rev-parse", "HEAD"),
        "ai-camp-schedule": git("rev-parse", "HEAD", cwd=ROOT / "ai-camp-schedule"),
    }
    deployed = None
    if os.environ.get("FORCE_PUBLISH") != "true":
        url = f"https://internal.aurange.cn/deployment.json?check={time.time_ns()}"
        try:
            with urllib.request.urlopen(url, timeout=20) as response:
                deployed = json.load(response)
        except (urllib.error.URLError, TimeoutError, ValueError):
            # A failed previous deployment must be retried, even if the gitlink
            # was already updated by that run.
            pass

    publish = deployed != versions
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as output:
            output.write(f"publish={str(publish).lower()}\n")
    if not publish:
        print("Published versions are current; no deployment needed.")
        return

    if SITE.exists():
        shutil.rmtree(SITE)
    SITE.mkdir()
    excluded = {".github", ".gitignore", ".gitmodules", "DEPLOYMENT.md", "ai-camp-schedule"}
    tracked = subprocess.check_output(["git", "ls-files", "-z"], cwd=ROOT).decode().split("\0")
    for name in filter(None, tracked):
        relative = Path(name)
        if relative.parts[0] in excluded:
            continue
        if relative.parts[0] == "camp-schedule" and (
            "tests" in relative.parts or relative.name in {"setup.sql", "package.json", "package-lock.json"}
        ):
            continue
        source = ROOT / relative
        if source.is_symlink() or not source.is_file():
            raise ValueError(f"Expected a regular published file: {name}")
        destination = SITE / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)

    # Copy static files only. Never execute code from the contributor repository.
    assets = ["index.html", "data.js"]
    # Supports publishing the portal before the independent AI repository is
    # upgraded; once either cloud asset exists, require the complete pair.
    if any((ROOT / "ai-camp-schedule" / name).exists() for name in ("config.js", "cloud.js")):
        assets += ["config.js", "cloud.js"]
    for name in assets:
        source = ROOT / "ai-camp-schedule" / name
        if source.is_symlink() or not source.is_file():
            raise ValueError(f"Expected an AI schedule static file: {name}")
        destination = SITE / "ai-camp-schedule" / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
    (SITE / ".nojekyll").touch()
    (SITE / "deployment.json").write_text(json.dumps(versions) + "\n")
    print(json.dumps(versions))


if __name__ == "__main__":
    assemble()
