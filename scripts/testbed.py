#!/usr/bin/env python3
import os
import sys
import argparse
from pathlib import Path

try:
    import yaml
except ImportError:
    print("Missing dependency: pyyaml. Install with: pip install pyyaml")
    sys.exit(1)

def mkdirp(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def write_file(path: Path, content: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")

def clean_dir(target: Path):
    """Remove contents of target safely."""
    if not target.exists():
        return
    for item in target.rglob("*"):
        try:
            if item.is_file():
                item.unlink()
        except Exception:
            pass
    # remove empty dirs
    for sub in sorted(list(target.glob("**/*")), reverse=True):
        if sub.is_dir() and not any(sub.iterdir()):
            try:
                sub.rmdir()
            except Exception:
                pass

def main():
    ap = argparse.ArgumentParser(description="Apply/reset pytenberg testbed from YAML.")
    ap.add_argument("--apply", metavar="testbed.yaml", required=True, help="Path to YAML config")
    ap.add_argument("--reset", action="store_true", help="Remove ignored outputs (out/, logs/, drop/*.msg)")
    args = ap.parse_args()

    root = Path.cwd()
    cfg_path = root / args.apply
    if not cfg_path.exists():
        print(f"Config not found: {cfg_path}")
        sys.exit(1)

    with open(cfg_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    paths = cfg.get("paths", {})
    drop = root / paths.get("drop", "./drop")
    out = root / paths.get("out", "./out")
    logs = root / paths.get("logs", "./logs")
    refs = root / paths.get("refs", "./refs")

    if args.reset:
        for target in [out, logs]:
            clean_dir(target)
        if drop.exists():
            for msg in drop.glob("*.msg"):
                try:
                    msg.unlink()
                except Exception:
                    pass

    # create directories
    for d in [drop, out, logs, refs]:
        mkdirp(d)

    # write refs templates
    for r in cfg.get("refs", []):
        name = r.get("name", "README_TEMPLATE.txt")
        content = r.get("content", "Template file for Pytenberg output folders")
        write_file(refs / name, content)

    # create SUBJECTS.txt in drop
    subjects = cfg.get("subjects", [])
    if subjects:
        text = "# Subject lines for .msg testing\n\n" + "\n".join(f"- {s}" for s in subjects)
        write_file(drop / "SUBJECTS.txt", text)

    active = cfg.get("active_pattern", "invoice")
    print("=" * 60)
    print("Pytenberg Testbed setup complete")
    print("=" * 60)
    print(f"Active pattern: {active}")
    print(f"Created folders: {drop}, {out}, {logs}, {refs}")
    print("\nNext steps:")
    print("  1) Create .msg emails using subjects in drop/SUBJECTS.txt")
    print("  2) Save them into the drop/ folder")
    print("  3) Set ACTIVE_PATTERN in pytenberg.py to match")
    print("  4) Run: python3 pytenberg.py")
    print("\nTo reset later, run:")
    print("  python3 scripts/testbed.py --apply testbed.yaml --reset")

if __name__ == "__main__":
    main()

