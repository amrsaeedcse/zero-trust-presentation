"""Master Presentation Compiler for Zero Trust Architecture (University Campus Networks).

Usage:
    python build.py                 # Build complete 9-member unified presentation
    python build.py --part 6        # Build only Part 6 into dist/part_06_preview.pptx
    python build.py --part 6 --validate   # Build + run OOXML validate.py automatically
    python build.py --list          # Show all 9 team parts and owned files
"""

from __future__ import annotations

import argparse
import importlib
import subprocess
import sys
from pathlib import Path

from core.theme import new_presentation

REPO_ROOT = Path(__file__).resolve().parent
DIST_DIR = REPO_ROOT / "dist"
VALIDATE_SCRIPT = REPO_ROOT / ".agents" / "skills" / "pptx" / "scripts" / "office" / "validate.py"

MAX_SLIDES_PER_PART = 7

PARTS_REGISTRY = {
    1: ("parts.part_01_threat_model", "build_part_01", "Ahmed Hassan Mohamed Mostafa", "Threat Model and Boundaries"),
    2: ("parts.part_02_identity_auth", "build_part_02", "Omar Abdelsalam Abdelfattah Mohamed Afifi", "Identity and Authentication"),
    3: ("parts.part_03_device_posture", "build_part_03", "Omar Mounir Mohamed Arafa", "Device Security Posture"),
    4: ("parts.part_04_authorization_policies", "build_part_04", "Omar Ahmed Abdelaziz Mashhour", "Authorization Policies"),
    5: ("parts.part_05_network_segmentation", "build_part_05", "Ali Mohamed Ali Abdallah Omara", "Network Segmentation"),
    6: ("parts.part_06_remote_access", "build_part_06", "Amr Abdelazeem Saeed Mohamed", "Remote Resource Access"),
    7: ("parts.part_07_monitoring_response", "build_part_07", "Omar Mohamed Fawzy Fahmy Morsi Aboelenein", "Monitoring and Response"),
    8: ("parts.part_08_adoption_evaluation", "build_part_08", "Amr Khaled Husseini Mohamed Tawfik", "Adoption and Evaluation"),
    9: ("parts.part_09_data_security", "build_part_09", "Mahmoud Mohamed Ragab Mohamed Hassan", "Data Security and Classification"),
}


def _run_part(prs, part_num: int) -> int:
    mod_name, fn_name, owner, topic = PARTS_REGISTRY[part_num]
    mod = importlib.import_module(mod_name)
    fn = getattr(mod, fn_name)
    before = len(prs.slides)
    returned = fn(prs)
    after = len(prs.slides)
    added = after - before
    if returned is not None and len(returned) != added:
        print(f"[WARN] Part {part_num:02d} returned {len(returned)} slides, actual added={added}")
    if added > MAX_SLIDES_PER_PART:
        raise ValueError(
            f"RULE 2 VIOLATION in Part {part_num:02d} ({owner}): "
            f"produced {added} slides (hard cap is {MAX_SLIDES_PER_PART} max; target 5-6)."
        )
    print(f"  Part {part_num:02d} | {added:2d} slide(s) | {topic} ({owner})")
    return added


def build_deck(part_only: int | None = None, out_path: Path | None = None) -> Path:
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    prs = new_presentation()

    if part_only is not None:
        if part_only not in PARTS_REGISTRY:
            raise SystemExit(f"Invalid --part {part_only}. Must be 1..9.")
        print(f"Building single-part preview: Part {part_only:02d}")
        _run_part(prs, part_only)
        target = out_path or (DIST_DIR / f"part_{part_only:02d}_preview.pptx")
    else:
        print("Building full unified deck (Intro + Parts 01-09 + Conclusion):")
        intro_mod = importlib.import_module("parts.part_00_intro")
        before = len(prs.slides)
        intro_mod.build_part_00(prs)
        print(f"  Part 00 | {len(prs.slides) - before:2d} slide(s) | Intro, Team Roster & Architecture Map")

        for num in sorted(PARTS_REGISTRY):
            _run_part(prs, num)

        outro_mod = importlib.import_module("parts.part_10_conclusion")
        before = len(prs.slides)
        outro_mod.build_part_10(prs)
        print(f"  Part 10 | {len(prs.slides) - before:2d} slide(s) | Conclusion & References")
        target = out_path or (DIST_DIR / "Zero_Trust_Campus_Networks.pptx")

    prs.save(str(target))
    print(f"\nSaved: {target} ({len(prs.slides)} total slide(s))")
    return target


def run_validation(pptx_path: Path) -> int:
    if not VALIDATE_SCRIPT.exists():
        print(f"[WARN] validate.py not found at {VALIDATE_SCRIPT}")
        return 1
    print(f"\nRunning OOXML validation on {pptx_path.name} ...")
    res = subprocess.run([sys.executable, str(VALIDATE_SCRIPT), str(pptx_path)])
    return res.returncode


def main() -> int:
    parser = argparse.ArgumentParser(description="Build Zero Trust Campus Networks PPTX presentation.")
    parser.add_argument("--part", type=int, choices=range(1, 10), metavar="1..9", help="Build only one member part")
    parser.add_argument("--out", type=Path, help="Optional custom output .pptx path")
    parser.add_argument("--validate", action="store_true", help="Run .agents/skills/pptx validate.py after build")
    parser.add_argument("--list", action="store_true", help="List all 9 parts and owners")
    args = parser.parse_args()

    if args.list:
        print("Zero Trust Architecture for University Campus Networks — Team Parts:")
        for num, (mod_name, _, owner, topic) in PARTS_REGISTRY.items():
            file_path = mod_name.replace(".", "/") + ".py"
            print(f"  Part {num:02d} | {file_path:40s} | {topic:34s} | {owner}")
        return 0

    out_file = build_deck(part_only=args.part, out_path=args.out)
    if args.validate:
        return run_validation(out_file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
