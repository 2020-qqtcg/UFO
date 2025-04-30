# Running UFO on SpreadsheetBench
This guide will help you run the UFO agent on the SpreadsheetBench benchmark.

## 1. Clone the SpreadsheetBench Repository
You can clone SpreadsheetBench anywhere.
```bash
git clone https://github.com/RUCKBReasoning/SpreadsheetBench.git
cd SpreadsheetBench
conda create -n SpreadsheetBench python=3.10
conda activate SpreadsheetBench
# Remove `vllm` from SpreadsheetBench requirements.txt
pip install -r requirements.txt
```

## Install UFO
Follow [UFO README](https://github.com/2020-qqtcg/UFO/blob/2020qq-tcg/ssb/README.md) install and configure UFO environment.

## 2. Prepare UFO Task Files
Step 1: Set Output Directory
Create the directory `outputs/custom_custom` inside your SpreadsheetBench dataset folder.

Then, open `process_scripts.py` and update the configuration section at the top of the file to match your environment:

```python
# --- Configuration ---

# PLEASE UPDATE THESE PATHS ACCORDING TO YOUR ENVIRONMENT

dataset_file_path = r"D:\code\SpreadsheetBench\data\sample_data_200\dataset.json"  # Path to your SpreadsheetBench dataset.json
input_base_dir = r"D:\code\SpreadsheetBench\data\sample_data_200"  # Directory containing spreadsheet folders (e.g., 'spreadsheet/59196')
output_dir = r"D:\code\SpreadsheetBench\data\sample_data_200\outputs\custom_custom"  # Output directory for the final .xlsx files
tasks_dir = r"D:\code\UFO\benchmark\tasks"  # Directory where UFO task files will be saved

# --- End Configuration ---
```
Step 2: Generate Task Files
Run the process_scripts.py script to generate UFO-compatible task files:

```bash
python benchmark/process_scripts.py
```
Task files will be saved under the benchmark/tasks directory.

## 3. Run UFO Agent
Execute the UFO agent in batch mode using the generated task files:

```bash
python -m ufo -m batch_normal -p [Path to benchmark/tasks]
```
Replace `[Path to benchmark/tasks]` with the actual path to the generated task directory.

## 4. Evaluate with SpreadsheetBench
Navigate to the evaluation folder inside SpreadsheetBench, then run the evaluation script:

```bash
cd evaluation
python evaluation.py --model custom --setting custom
```