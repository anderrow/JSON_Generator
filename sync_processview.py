import argparse
import filecmp
import os
import shutil
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent
DEFAULT_SOURCE_DIR = ROOT_DIR / "SOURCE FILES"
DEFAULT_JSON_DIR = ROOT_DIR / "JSON FILES"
DEFAULT_TRANSLATIONS_SUBDIR = Path("hmi") / "translations"
FALLBACK_LANGUAGE_FILE = "en.json"
SUPPORTED_LANGUAGE_FILES = {
    "bg.json",
    "da.json",
    "de.json",
    "en.json",
    "fr.json",
    "nl.json",
    "uk.json",
}

PROJECT_TARGETS = {
    "AVA": "ava",
    "CAR": "crevin",
    "DSM": "dsm",
    "UFA": "ufa",
    "VILO": "vilo",
}


@dataclass(frozen=True)
class ExportBundle:
    bundle_name: str
    project_name: str
    source_csv: Path
    json_dir: Path
    target_dir: Path
    target_rel: Path


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate JSON files and sync them into the ProcessView translations repo."
    )
    parser.add_argument(
        "--processview-repo",
        help="Path to the local processview git repository. Defaults to PROCESSVIEW_REPO or common clone locations.",
    )
    parser.add_argument(
        "--source-dir",
        default=str(DEFAULT_SOURCE_DIR),
        help="Directory that contains the source CSV files.",
    )
    parser.add_argument(
        "--json-dir",
        default=str(DEFAULT_JSON_DIR),
        help="Directory that contains the generated JSON folders.",
    )
    parser.add_argument(
        "--translations-subdir",
        default=str(DEFAULT_TRANSLATIONS_SUBDIR),
        help="Path inside the processview repo where translation folders live.",
    )
    parser.add_argument(
        "--project",
        action="append",
        help="Restrict the sync to one project code. Repeat the flag for multiple projects.",
    )
    parser.add_argument(
        "--skip-generate",
        action="store_true",
        help="Skip running JsonGenerator.py before syncing.",
    )
    parser.add_argument(
        "--skip-pull",
        action="store_true",
        help="Skip git pull --ff-only before copying translations.",
    )
    parser.add_argument(
        "--push",
        action="store_true",
        help="Stage, commit and push the updated translation folders after syncing.",
    )
    parser.add_argument(
        "--commit-message",
        help="Commit message to use together with --push.",
    )
    parser.add_argument(
        "--allow-dirty",
        action="store_true",
        help="Allow copying into a repo that already has local changes. Pull is still blocked on a dirty repo.",
    )
    parser.add_argument(
        "--prune",
        action="store_true",
        help="Remove supported language files from the target folder when they are missing in the generated JSON folder.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview the sync without running the generator, pulling, copying or pushing.",
    )
    return parser.parse_args()


def resolve_processview_repo(cli_value):
    candidates = []

    if cli_value:
        candidates.append(Path(cli_value))

    env_value = os.environ.get("PROCESSVIEW_REPO")
    if env_value:
        candidates.append(Path(env_value))

    home = Path.home()
    candidates.extend(
        [
            home / "Documents" / "GitHub" / "processview",
            home / "Documents" / "Git" / "processview",
            home / "Documents" / "Git" / "Proccesview",
        ]
    )

    for candidate in candidates:
        repo_path = candidate.expanduser().resolve()
        if (repo_path / ".git").exists():
            return repo_path

    searched = "\n".join(f" - {candidate.expanduser()}" for candidate in candidates)
    raise RuntimeError(
        "Could not locate the processview repository.\n"
        "Use --processview-repo or set PROCESSVIEW_REPO.\n"
        f"Searched:\n{searched}"
    )


def run_command(command, cwd=None, dry_run=False, env=None):
    printable = " ".join(f'"{part}"' if " " in str(part) else str(part) for part in command)
    if dry_run:
        print(f"[dry-run] {printable}")
        return ""

    completed = subprocess.run(
        command,
        cwd=cwd,
        capture_output=True,
        text=True,
        env=env,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"Command failed ({completed.returncode}): {printable}\n"
            f"stdout:\n{completed.stdout}\n"
            f"stderr:\n{completed.stderr}"
        )
    return completed.stdout.strip()


def get_repo_status(repo_path):
    return run_command(["git", "-C", str(repo_path), "status", "--porcelain"])


def ensure_clean_repo(repo_path, allow_dirty, skip_pull):
    status = get_repo_status(repo_path)
    if not status:
        return

    if not skip_pull:
        raise RuntimeError(
            "The processview repo has local changes, so git pull --ff-only would be unsafe.\n"
            "Commit or stash those changes, or rerun with --skip-pull after reviewing them."
        )

    if not allow_dirty:
        raise RuntimeError(
            "The processview repo has local changes.\n"
            "Rerun with --allow-dirty if you intentionally want to copy only the translation updates."
        )


def discover_bundles(source_dir, json_dir, repo_path, translations_subdir, selected_projects):
    grouped = defaultdict(list)

    for source_csv in sorted(source_dir.glob("*.csv")):
        bundle_name = source_csv.stem
        project_name = bundle_name.split("_", 1)[0].upper()

        if selected_projects and project_name not in selected_projects:
            continue

        target_name = PROJECT_TARGETS.get(project_name)
        if not target_name:
            print(f"Skipping {bundle_name}: no target folder mapping for project {project_name}.")
            continue

        target_rel = translations_subdir / target_name
        bundle = ExportBundle(
            bundle_name=bundle_name,
            project_name=project_name,
            source_csv=source_csv,
            json_dir=json_dir / bundle_name,
            target_dir=repo_path / target_rel,
            target_rel=target_rel,
        )
        grouped[project_name].append(bundle)

    duplicates = {project: bundles for project, bundles in grouped.items() if len(bundles) > 1}
    if duplicates:
        lines = ["Found multiple CSV files for the same project. Keep only one active export per project:"]
        for project, bundles in sorted(duplicates.items()):
            names = ", ".join(bundle.bundle_name for bundle in bundles)
            lines.append(f" - {project}: {names}")
        raise RuntimeError("\n".join(lines))

    bundles = sorted(
        (bundle_list[0] for bundle_list in grouped.values()),
        key=lambda bundle: (bundle.project_name, bundle.bundle_name),
    )

    if selected_projects:
        found_projects = {bundle.project_name for bundle in bundles}
        missing_projects = sorted(selected_projects - found_projects)
        if missing_projects:
            raise RuntimeError(
                "No active CSV export found for: " + ", ".join(missing_projects)
            )

    if not bundles:
        raise RuntimeError(f"No CSV files found in {source_dir}.")

    return bundles


def build_copy_plan(source_dir, target_dir, prune):
    all_source_files = {
        path.name: path for path in sorted(source_dir.glob("*.json")) if path.is_file()
    }
    source_files = {
        filename: path
        for filename, path in all_source_files.items()
        if filename in SUPPORTED_LANGUAGE_FILES
    }

    if not source_files:
        raise RuntimeError(f"No JSON files found in {source_dir}.")

    ignored_source_files = sorted(set(all_source_files) - set(source_files))
    effective_source_files = dict(source_files)
    fallback_files = []

    if FALLBACK_LANGUAGE_FILE in source_files:
        fallback_source = source_files[FALLBACK_LANGUAGE_FILE]
        for filename in sorted(SUPPORTED_LANGUAGE_FILES):
            if filename not in effective_source_files:
                effective_source_files[filename] = fallback_source
                fallback_files.append(filename)

    target_files = {
        path.name: path
        for path in sorted(target_dir.glob("*.json"))
        if path.is_file() and path.name in SUPPORTED_LANGUAGE_FILES
    }

    files_to_copy = []
    unchanged_files = []
    files_to_remove = (
        sorted(target_files.keys() - effective_source_files.keys()) if prune else []
    )

    for filename, source_file in sorted(effective_source_files.items()):
        target_file = target_dir / filename
        if target_file.exists() and filecmp.cmp(source_file, target_file, shallow=False):
            unchanged_files.append(filename)
            continue
        files_to_copy.append(filename)

    return files_to_copy, files_to_remove, unchanged_files, ignored_source_files, fallback_files, effective_source_files


def sync_bundle(bundle, dry_run, prune):
    if not bundle.json_dir.exists():
        raise RuntimeError(
            f"Generated JSON folder not found for {bundle.bundle_name}: {bundle.json_dir}"
        )

    bundle.target_dir.mkdir(parents=True, exist_ok=True)

    (
        files_to_copy,
        files_to_remove,
        unchanged_files,
        ignored_source_files,
        fallback_files,
        effective_source_files,
    ) = build_copy_plan(
        bundle.json_dir, bundle.target_dir, prune=prune
    )

    print(f"{bundle.bundle_name} -> {bundle.target_rel.as_posix()}")
    print(
        f"  copy: {len(files_to_copy)} file(s), "
        f"remove: {len(files_to_remove)} file(s), "
        f"unchanged: {len(unchanged_files)} file(s)"
    )
    if ignored_source_files:
        print(f"  ignored unsupported language file(s): {', '.join(ignored_source_files)}")
    if fallback_files:
        print(f"  fallback from {FALLBACK_LANGUAGE_FILE}: {', '.join(fallback_files)}")

    if dry_run:
        for filename in files_to_copy:
            source_label = effective_source_files[filename].name
            if source_label == filename:
                print(f"  [dry-run] copy {filename}")
            else:
                print(f"  [dry-run] copy {filename} from {source_label}")
        for filename in files_to_remove:
            print(f"  [dry-run] remove {filename}")
        return bool(files_to_copy or files_to_remove)

    for filename in files_to_remove:
        (bundle.target_dir / filename).unlink()

    for filename in files_to_copy:
        shutil.copy2(effective_source_files[filename], bundle.target_dir / filename)

    return bool(files_to_copy or files_to_remove)


def build_commit_message(bundles):
    names = ", ".join(bundle.bundle_name for bundle in bundles)
    return f"Update ProcessView translations: {names}"


def stage_commit_and_push(repo_path, bundles, commit_message, dry_run):
    target_paths = sorted({bundle.target_rel.as_posix() for bundle in bundles})
    status = run_command(
        ["git", "-C", str(repo_path), "status", "--porcelain", "--", *target_paths]
    )

    if not status:
        print("No git changes detected in the translation folders.")
        return

    if dry_run:
        print("[dry-run] git add/commit/push would run for:")
        for target_path in target_paths:
            print(f"  - {target_path}")
        return

    run_command(["git", "-C", str(repo_path), "add", "--", *target_paths])

    staged = run_command(
        ["git", "-C", str(repo_path), "diff", "--cached", "--name-only", "--", *target_paths]
    )
    if not staged:
        print("Nothing was staged after git add; skipping commit.")
        return

    final_message = commit_message or build_commit_message(bundles)
    run_command(["git", "-C", str(repo_path), "commit", "-m", final_message])
    run_command(["git", "-C", str(repo_path), "push"])
    print("Changes committed and pushed.")


def run_generator(dry_run, source_dir, json_dir):
    generator_path = SCRIPT_DIR / "JsonGenerator.py"
    env = os.environ.copy()
    env["JSON_GENERATOR_SOURCE_DIR"] = str(source_dir)
    env["JSON_GENERATOR_OUTPUT_DIR"] = str(json_dir)
    if dry_run:
        print(
            f"[dry-run] JSON_GENERATOR_SOURCE_DIR={source_dir} "
            f"JSON_GENERATOR_OUTPUT_DIR={json_dir} "
            f"{sys.executable} {generator_path}"
        )
        return

    print("Running JsonGenerator.py...")
    run_command([sys.executable, str(generator_path)], cwd=SCRIPT_DIR, env=env)


def main():
    args = parse_args()

    source_dir = Path(args.source_dir).expanduser().resolve()
    json_dir = Path(args.json_dir).expanduser().resolve()
    translations_subdir = Path(args.translations_subdir)
    selected_projects = {project.upper() for project in args.project or []}

    if not source_dir.exists():
        raise RuntimeError(f"Source directory not found: {source_dir}")

    if not json_dir.exists() and (args.skip_generate or args.dry_run):
        raise RuntimeError(f"JSON directory not found: {json_dir}")

    repo_path = resolve_processview_repo(args.processview_repo)

    print(f"ProcessView repo: {repo_path}")
    print(f"Source CSV dir:   {source_dir}")
    print(f"JSON output dir:  {json_dir}")

    if args.dry_run:
        print("Dry-run mode: generation, pull, copy and push will be skipped.")

    ensure_clean_repo(
        repo_path=repo_path,
        allow_dirty=args.allow_dirty,
        skip_pull=args.skip_pull or args.dry_run,
    )

    if not args.skip_generate:
        run_generator(dry_run=args.dry_run, source_dir=source_dir, json_dir=json_dir)

    bundles = discover_bundles(
        source_dir=source_dir,
        json_dir=json_dir,
        repo_path=repo_path,
        translations_subdir=translations_subdir,
        selected_projects=selected_projects,
    )

    print("Bundles to sync:")
    for bundle in bundles:
        print(f" - {bundle.bundle_name} ({bundle.project_name})")

    if not args.skip_pull and not args.dry_run:
        print("Running git pull --ff-only...")
        run_command(["git", "-C", str(repo_path), "pull", "--ff-only"])
    elif args.dry_run and not args.skip_pull:
        print("[dry-run] git pull --ff-only")

    changed_bundles = []
    for bundle in bundles:
        if sync_bundle(bundle, dry_run=args.dry_run, prune=args.prune):
            changed_bundles.append(bundle)

    if not changed_bundles:
        print("All translation folders are already up to date.")
        return

    if args.push:
        stage_commit_and_push(
            repo_path=repo_path,
            bundles=changed_bundles,
            commit_message=args.commit_message,
            dry_run=args.dry_run,
        )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
