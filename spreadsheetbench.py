"""SpreadsheetBench — OpenReward sandbox environment for spreadsheet manipulation.

912 real-world Excel forum questions with OJ-style evaluation. Agents explore
input spreadsheets, write Python scripts to perform the required manipulation,
and submit for automated testing across multiple test cases.

Paper: https://arxiv.org/abs/2406.14991
Dataset: https://huggingface.co/datasets/KAKA22/SpreadsheetBench
"""

import json
import logging
import os
from pathlib import Path

from openreward import AsyncOpenReward, SandboxBucketConfig, SandboxSettings
from openreward.environments import Environment, JSONObject, TextBlock, ToolOutput, tool
from openreward.toolsets import ExcelToolset
from pydantic import BaseModel

from evaluate import compare_workbooks

logger = logging.getLogger(__name__)

# --- Module-level data loading ---

if os.path.exists("/orwd_data"):
    _DATA_DIR = Path("/orwd_data/server_data/spreadsheetbench")
else:
    _DATA_DIR = Path(__file__).parent / "server_data" / "spreadsheetbench"

_all_records: dict[str, dict] = {}
_tasks: list[JSONObject] = []

_dataset_path = _DATA_DIR / "dataset.json"
if not _dataset_path.exists():
    logger.warning(f"Data file not found: {_dataset_path}")
else:
    with open(_dataset_path) as _f:
        _dataset = json.load(_f)
    for _record in _dataset:
        _record_id = str(_record["id"])
        _all_records[_record_id] = _record
        _tasks.append({
            "id": _record_id,
            "instruction_type": _record["instruction_type"],
        })


# --- Pydantic parameter models ---

class BashParams(BaseModel, extra="forbid"):
    command: str


class SubmitParams(BaseModel, extra="forbid"):
    """Submit a Python script for OJ-style evaluation."""
    script_path: str


# --- Environment class ---

class SpreadsheetBench(Environment):
    toolsets = [ExcelToolset]

    def __init__(self, task_spec: JSONObject, secrets: dict[str, str] = {}) -> None:
        super().__init__(task_spec)

        record_id = str(task_spec["id"])
        if record_id not in _all_records:
            raise ValueError(f"Unknown task id: {record_id}")

        record = _all_records[record_id]
        self.task_id: str = record_id
        self.instruction: str = record["instruction"]
        self.instruction_type: str = record["instruction_type"]
        self.answer_position: str = record["answer_position"]
        self.num_test_cases: int = record["num_test_cases"]
        self.answer_sheet: str = record.get("answer_sheet", "")
        self.data_position: str = record.get("data_position", "")

        # Build list of input filenames (using the original task ID for file naming)
        self.original_id: str = str(record["id"])
        self.input_files: list[str] = [
            f"{i}_{self.original_id}_input.xlsx"
            for i in range(1, self.num_test_cases + 1)
        ]

        api_key = (
            secrets.get("OPENREWARD_API_KEY")
            or secrets.get("api_key")
            or os.environ.get("OPENREWARD_API_KEY", "").strip('"')
        )
        if not api_key:
            raise ValueError("OpenReward API key is required (pass as OPENREWARD_API_KEY)")

        self.sandbox_settings = SandboxSettings(
            environment="GeneralReasoning/SpreadsheetBench",
            image="generalreasoning/knowledge-worker:latest",
            machine_size="0.5:1",
            block_network=False,
            bucket_config=SandboxBucketConfig(
                mount_path="/data",
                read_only=True,
                only_dir=f"spreadsheetbench/inputs/{self.task_id}",
            ),
        )

        or_client = AsyncOpenReward(api_key=api_key)
        self.sandbox = or_client.sandbox(self.sandbox_settings)
        self.submitted = False

    async def setup(self) -> None:
        await self.sandbox.start()

    async def teardown(self) -> None:
        await self.sandbox.stop()

    @classmethod
    def list_splits(cls) -> list[str]:
        return ["test"]

    @classmethod
    def list_tasks(cls, split: str) -> list[JSONObject]:
        if split == "test":
            return _tasks
        return []

    def get_prompt(self) -> list[TextBlock]:
        file_list = "\n".join(f"  - /data/{f}" for f in self.input_files)

        prompt = f"""You are solving a SpreadsheetBench task — a real-world spreadsheet manipulation problem.

## Task

{self.instruction}

## Spreadsheet Details

- **Instruction type**: {self.instruction_type}
- **Answer position**: {self.answer_position} — the cell(s) where your result must appear
- **Data position**: {self.data_position or "See spreadsheet"}
- **Answer sheet**: {self.answer_sheet or "See answer position or use the first sheet"}

## Available Input Files

The following input spreadsheet(s) are available (read-only):
{file_list}

These files have the same structure but different data values. Your solution must work correctly on all of them.

## How to Submit

1. Explore the input spreadsheets using the Excel tools (e.g., `excel_read_tab`, `excel_list_tabs_in_spreadsheet`) or `bash`.
2. Write a Python script that:
   - Takes an **input file path** as the first argument (`sys.argv[1]`)
   - Takes an **output file path** as the second argument (`sys.argv[2]`)
   - Reads the input spreadsheet, performs the required manipulation
   - Saves the result to the output path
3. Test your script on the available input files.
4. Call `submit` with the path to your script.

Your script will be run on each input file independently and evaluated against expected answers.

## Important Notes

- Write **computed values** into cells, not Excel formulas. The grader reads cell values directly.
- The `openpyxl` library is available in the environment.
- Only the cells at the **answer position** are checked — you don't need to modify other cells.
- All {self.num_test_cases} test case(s) must pass for reward=1.0.

## Example Script Structure

```python
import sys
import openpyxl

input_path = sys.argv[1]
output_path = sys.argv[2]

wb = openpyxl.load_workbook(input_path)
ws = wb.active  # or wb["SheetName"]

# ... perform manipulation ...

wb.save(output_path)
```"""

        return [TextBlock(text=prompt)]

    @tool
    async def bash(self, params: BashParams) -> ToolOutput:
        """Execute a bash command in the sandbox environment."""
        result = await self.sandbox.run(params.command.strip())
        output = result.output
        code = result.return_code

        if result.truncated:
            output = f"...(truncated, output exceeded limit)\n{output}"

        return ToolOutput(
            blocks=[TextBlock(text=f"{output}\n\n(exit {code})")],
            metadata={"output": output, "exit_code": code, "truncated": result.truncated},
            reward=0.0,
            finished=False,
        )

    @tool
    async def submit(self, params: SubmitParams) -> ToolOutput:
        """Submit a Python script for OJ-style evaluation.

        The script is run on each test case input file. All test cases must
        produce correct output at the answer position for reward=1.0.
        This is a terminal action — you get one submission attempt.
        """
        if self.submitted:
            return ToolOutput(
                blocks=[TextBlock(text="Already submitted. Only one submission is allowed.")],
                metadata={"error": "already_submitted"},
                reward=0.0,
                finished=True,
            )

        self.submitted = True
        script_path = params.script_path

        # Verify script exists
        check_result = await self.sandbox.run(f"test -f '{script_path}' && echo EXISTS")
        if "EXISTS" not in check_result.output:
            return ToolOutput(
                blocks=[TextBlock(text=f"Error: Script not found at {script_path}")],
                metadata={"error": "script_not_found"},
                reward=0.0,
                finished=True,
            )

        results = []
        details = []
        answers_dir = _DATA_DIR / "answers" / self.task_id

        for i in range(1, self.num_test_cases + 1):
            input_file = f"/data/{i}_{self.original_id}_input.xlsx"
            output_file = f"/tmp/output_{i}.xlsx"

            # Run the agent's script
            run_result = await self.sandbox.run(
                f"python '{script_path}' '{input_file}' '{output_file}' 2>&1",
                timeout=120,
            )
            run_output = run_result.output
            exit_code = run_result.return_code

            if exit_code != 0:
                results.append(False)
                error_msg = run_output[:500] if run_output else "No output"
                details.append(f"Test case {i}: FAIL — script error (exit {exit_code}): {error_msg}")
                continue

            # Download the output file from sandbox
            try:
                output_bytes = await self.sandbox.download(output_file)
            except Exception as e:
                results.append(False)
                details.append(f"Test case {i}: FAIL — output file not found at {output_file}")
                continue

            # Compare against answer file
            answer_path = answers_dir / f"{i}_{self.original_id}_answer.xlsx"
            if not answer_path.exists():
                results.append(False)
                details.append(f"Test case {i}: FAIL — answer file missing (server error)")
                logger.error(f"Answer file not found: {answer_path}")
                continue

            try:
                passed = compare_workbooks(answer_path, output_bytes, self.answer_position)
                results.append(passed)
                details.append(f"Test case {i}: {'PASS' if passed else 'FAIL — cell values do not match'}")
            except Exception as e:
                results.append(False)
                details.append(f"Test case {i}: FAIL — comparison error: {e}")
                logger.exception(f"Comparison error for task {self.task_id}, test case {i}")

        num_passed = sum(results)
        num_total = len(results)
        all_pass = num_passed == num_total and num_total > 0
        reward = 1.0 if all_pass else 0.0

        summary = "\n".join(details)
        result_text = f"""Submission Results:
- Passed: {num_passed}/{num_total} test cases
- Reward: {reward}

{summary}"""

        return ToolOutput(
            blocks=[TextBlock(text=result_text)],
            metadata={
                "num_passed": num_passed,
                "num_total": num_total,
                "results": results,
            },
            reward=reward,
            finished=True,
        )
