# SpreadsheetBench

[![OpenReward Environment](https://img.shields.io/badge/%E2%AD%90%20OpenReward-Environment-f7e6cc)](https://openreward.ai/GeneralReasoning/SpreadsheetBench)

## Description

**SpreadsheetBench** is an environment for evaluating LLM agents on real-world spreadsheet manipulation tasks. It comprises 905 instructions (from 912 original, 7 excluded due to broken metadata) sourced from online Excel forums, each with multiple test cases (typically 3) to ensure solution generality via OJ-style evaluation.

## Capabilities

- Reading and analyzing Excel spreadsheet structure
- Writing Python code to manipulate spreadsheet data
- Handling cell-level and sheet-level operations
- Producing general solutions that work across multiple test cases

## Compute Requirements

- Sandbox: 0.5 CPU / 1 GB memory per session
- Network access: enabled (not blocked)
- No GPU required

## License

[CC-BY-SA-4.0](https://creativecommons.org/licenses/by-sa/4.0/)

## Tasks

Single `test` split with 905 tasks. Each task has 1–3 test cases (~2,700 total) with different spreadsheet values but the same instruction. Tasks are divided into:

- **Cell-Level Manipulation** (560 tasks): modifying specific cells or ranges
- **Sheet-Level Manipulation** (345 tasks): modifying entire sheets, cross-sheet operations

## Reward Structure

Binary reward (0.0 or 1.0). The agent's Python script is executed on each test case input file. Cell values at the specified `answer_position` are compared against ground-truth answer files. All test cases must pass for reward=1.0 (OJ-style hard metric).

## Data

- **Source**: [KAKA22/SpreadsheetBench](https://huggingface.co/datasets/KAKA22/SpreadsheetBench) on HuggingFace
- **Format**: Excel `.xlsx` files with JSON metadata
- **Size**: ~91 MB compressed
- Input spreadsheets mounted read-only at `/data/` in the sandbox

## Tools

- `bash` — Execute bash commands in the sandbox for writing code and testing
- `submit` — Submit a Python script for OJ-style evaluation across all test cases
- `excel_list_tabs_in_spreadsheet` — List all worksheet names
- `excel_read_tab` — Read data from a specific worksheet
- `excel_read_csv` — Read CSV files
- `excel_create_spreadsheet`, `excel_add_tab`, `excel_edit_spreadsheet`, `excel_add_content_text`, `excel_delete_content_cell`, `excel_create_chart`, `excel_delete_tab`, `excel_delete_spreadsheet` — Full Excel manipulation via the ExcelToolset

## Time Horizon

Multi-turn. Agents typically explore the spreadsheet, write a solution script, test it, and submit. Average interaction involves 5–15 tool calls.

## Environment Difficulty

The original benchmark reports ChatGPT Agent achieving 45.5% task success rate, indicating substantial difficulty. Tasks range from simple cell extraction to complex multi-sheet operations.

## Safety

Tasks involve spreadsheet data manipulation only. Input data is sourced from public Excel forum questions. No personally identifiable information or sensitive data.

## Citations

```bibtex
@inproceedings{ma2024spreadsheetbench,
  title={SpreadsheetBench: Towards Challenging Real World Spreadsheet Manipulation},
  author={Ma, Zeyao and Zhang, Bohan and Zhang, Jing and Yu, Jifan and Zhang, Xiaokang and Zhang, Xiaohan and Luo, Sijia and Wang, Xi and Tang, Jie},
  booktitle={Advances in Neural Information Processing Systems},
  year={2024}
}
```
