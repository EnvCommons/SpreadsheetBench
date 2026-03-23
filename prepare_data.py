"""Download and prepare SpreadsheetBench data for the OpenReward environment.

Downloads the KAKA22/SpreadsheetBench dataset from HuggingFace, extracts the
archive, and organizes files into two directory trees:
  - bucket_data/spreadsheetbench/inputs/{id}/  — input xlsx files (for sandbox)
  - server_data/spreadsheetbench/             — dataset.json + answer xlsx files

Tasks with broken metadata (malformed answer_position, wrong sheet names) are
detected via a self-compare validation step and excluded automatically.

Usage:
    uv run python prepare_data.py
"""

import json
import shutil
import tarfile
from pathlib import Path

from huggingface_hub import hf_hub_download


REPO_ID = "KAKA22/SpreadsheetBench"
ARCHIVE_NAME = "spreadsheetbench_912_v0.1.tar.gz"
EXTRACT_ROOT = "all_data_912_v0.1"


def count_test_cases(spreadsheet_dir: Path, task_id: str) -> int:
    """Count the number of test cases for a task by scanning input files."""
    count = 0
    for i in range(1, 100):
        if (spreadsheet_dir / f"{i}_{task_id}_input.xlsx").exists():
            count += 1
        else:
            break
    return count


def validate_task(answer_path: Path, answer_position: str) -> tuple[bool, str]:
    """Validate a task by self-comparing its first answer file.

    Loads the answer xlsx and compares it against itself at the given
    answer_position. If this fails, the task's metadata is broken (wrong
    sheet names, malformed cell references, etc.) and it should be excluded.

    Returns (passed, reason).
    """
    from evaluate import compare_workbooks

    try:
        result = compare_workbooks(answer_path, answer_path, answer_position)
        if result:
            return True, ""
        else:
            return False, "self-compare mismatch (likely wrong sheet name in metadata)"
    except Exception as e:
        return False, str(e)


def main():
    output_dir = Path(__file__).parent

    # Step 1: Download from HuggingFace
    print(f"Downloading {ARCHIVE_NAME} from {REPO_ID}...")
    archive_path = hf_hub_download(
        repo_id=REPO_ID,
        filename=ARCHIVE_NAME,
        repo_type="dataset",
        local_dir=str(output_dir),
    )
    print(f"Downloaded to {archive_path}")

    # Step 2: Extract
    extract_dir = output_dir / EXTRACT_ROOT
    if extract_dir.exists():
        print(f"Removing existing {extract_dir}...")
        shutil.rmtree(extract_dir)

    print("Extracting archive...")
    with tarfile.open(archive_path, "r:gz") as tar:
        tar.extractall(path=str(output_dir), filter="data")
    print(f"Extracted to {extract_dir}")

    # Step 3: Load dataset.json
    dataset_path = extract_dir / "dataset.json"
    with open(dataset_path) as f:
        dataset = json.load(f)
    print(f"Loaded {len(dataset)} tasks from dataset.json")

    # Step 4: Organize into bucket_data and server_data
    bucket_dir = output_dir / "bucket_data" / "spreadsheetbench" / "inputs"
    server_dir = output_dir / "server_data" / "spreadsheetbench"
    answers_dir = server_dir / "answers"

    for d in [bucket_dir, answers_dir]:
        d.mkdir(parents=True, exist_ok=True)

    # First pass: copy files and build candidate task list
    candidates = []
    total_input_files = 0
    total_answer_files = 0

    for task in dataset:
        task_id = str(task["id"])
        spreadsheet_dir = extract_dir / "spreadsheet" / str(task["id"])

        if not spreadsheet_dir.exists():
            print(f"  WARNING: Missing spreadsheet dir for task {task_id}")
            continue

        # Count test cases dynamically
        num_test_cases = count_test_cases(spreadsheet_dir, str(task["id"]))
        if num_test_cases == 0:
            print(f"  WARNING: No test cases found for task {task_id}")
            continue

        # Copy input files to bucket_data
        task_input_dir = bucket_dir / task_id
        task_input_dir.mkdir(parents=True, exist_ok=True)

        for i in range(1, num_test_cases + 1):
            input_file = spreadsheet_dir / f"{i}_{task['id']}_input.xlsx"
            if input_file.exists():
                shutil.copy2(input_file, task_input_dir / input_file.name)
                total_input_files += 1

        # Copy answer files to server_data
        task_answer_dir = answers_dir / task_id
        task_answer_dir.mkdir(parents=True, exist_ok=True)

        for i in range(1, num_test_cases + 1):
            answer_file = spreadsheet_dir / f"{i}_{task['id']}_answer.xlsx"
            if answer_file.exists():
                shutil.copy2(answer_file, task_answer_dir / answer_file.name)
                total_answer_files += 1

        # Build enriched task record
        enriched_task = {
            "id": task_id,
            "instruction": task["instruction"],
            "instruction_type": task["instruction_type"],
            "answer_position": task["answer_position"],
            "num_test_cases": num_test_cases,
        }
        # Include optional fields if present
        if "answer_sheet" in task:
            enriched_task["answer_sheet"] = task["answer_sheet"]
        if "data_position" in task:
            enriched_task["data_position"] = task["data_position"]

        candidates.append(enriched_task)

    # Step 5: Validate — self-compare each task's first answer file
    # Tasks with broken metadata (wrong sheet names, malformed cell refs) are excluded.
    print(f"\nValidating {len(candidates)} tasks (self-compare answer files)...")
    enriched_tasks = []
    excluded = []

    for task in candidates:
        task_id = task["id"]
        original_id = task["id"]  # same as task_id since we already stringified
        answer_path = answers_dir / task_id / f"1_{original_id}_answer.xlsx"

        if not answer_path.exists():
            excluded.append((task_id, "answer file missing"))
            continue

        passed, reason = validate_task(answer_path, task["answer_position"])
        if passed:
            enriched_tasks.append(task)
        else:
            excluded.append((task_id, reason))

    # Sort for stable ordering
    enriched_tasks.sort(key=lambda t: t["id"])

    # Save enriched dataset.json
    enriched_path = server_dir / "dataset.json"
    with open(enriched_path, "w") as f:
        json.dump(enriched_tasks, f, indent=2)

    # Clean up excluded tasks' files from bucket and server data
    for task_id, _ in excluded:
        task_input_dir = bucket_dir / task_id
        task_answer_dir = answers_dir / task_id
        if task_input_dir.exists():
            shutil.rmtree(task_input_dir)
        if task_answer_dir.exists():
            shutil.rmtree(task_answer_dir)

    print(f"\nDone!")
    print(f"  Tasks: {len(enriched_tasks)} (excluded {len(excluded)} with broken metadata)")
    print(f"  Input files: {total_input_files}")
    print(f"  Answer files: {total_answer_files}")
    print(f"  Bucket data: {bucket_dir}")
    print(f"  Server data: {server_dir}")

    if excluded:
        print(f"\n  Excluded tasks:")
        for task_id, reason in excluded:
            print(f"    {task_id}: {reason}")

    # Instruction type distribution
    cell_level = sum(1 for t in enriched_tasks if t["instruction_type"] == "Cell-Level Manipulation")
    sheet_level = sum(1 for t in enriched_tasks if t["instruction_type"] == "Sheet-Level Manipulation")
    print(f"\n  Cell-Level: {cell_level}, Sheet-Level: {sheet_level}")

    print(f"\nNext steps:")
    print(f"  1. Upload bucket data:  openreward bucket upload {bucket_dir.parent}/ spreadsheetbench/")
    print(f"  2. Deploy server with server_data/ mounted at /orwd_data/")


if __name__ == "__main__":
    main()
