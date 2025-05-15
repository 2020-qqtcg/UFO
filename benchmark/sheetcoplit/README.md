# Running UFO on SheetCoplit
This guide will help you run the UFO agent on the SheetCoplit benchmark.

## 1. Clone the SheetCoplit Repository
You can clone SheetCoplit anywhere.
```bash
git clone https://github.com/RUCKBReasoning/SpreadsheetBench.git
cd SpreadsheetBench
conda create -n SpreadsheetBench python=3.10
conda activate SpreadsheetBench
# Remove `vllm` from SpreadsheetBench requirements.txt
pip install -r requirements.txt
```

## Install UFO
### 🛠️ Step 1: Installation
UFO requires **Python >= 3.10** running on **Windows OS >= 10**. It can be installed by running the following command:
```powershell
# [optional to create conda environment]
# conda create -n ufo python=3.10
# conda activate ufo

# clone the repository
git clone https://github.com/2020-qqtcg/UFO
cd UFO
git checkout 2020qq-tcg/ssb
# install the requirements
pip install -r requirements.txt
# If you want to use the Qwen as your LLMs, uncomment the related libs.
```

### ⚙️ Step 2: Configure the LLMs
Before running UFO, you need to provide your LLM configurations **individually for HostAgent and AppAgent**. You can create your own config file `ufo/config/config.yaml`, by copying the `ufo/config/config.yaml.template` and editing config for **HOST_AGENT** and **APP_AGENT** as follows: 

```powershell
copy ufo\config\config.yaml.template ufo\config\config.yaml
notepad ufo\config\config.yaml   # paste your key & endpoint
```

#### OpenAI
```yaml
VISUAL_MODE: True, # Whether to use the visual mode
API_TYPE: "openai" , # The API type, "openai" for the OpenAI API.  
API_BASE: "https://api.openai.com/v1/chat/completions", # The the OpenAI API endpoint.
API_KEY: "sk-",  # The OpenAI API key, begin with sk-
API_VERSION: "2024-02-15-preview", # "2024-02-15-preview" by default
API_MODEL: "gpt-4o",  # The only OpenAI model
```

#### Azure OpenAI (AOAI)
```yaml
VISUAL_MODE: True, # Whether to use the visual mode
API_TYPE: "aoai" , # The API type, "aoai" for the Azure OpenAI.  
API_BASE: "YOUR_ENDPOINT", #  The AOAI API address. Format: https://{your-resource-name}.openai.azure.com
API_KEY: "YOUR_KEY",  # The aoai API key
API_VERSION: "2024-02-15-preview", # "2024-02-15-preview" by default
API_MODEL: "gpt-4o",  # The only OpenAI model
API_DEPLOYMENT_ID: "YOUR_AOAI_DEPLOYMENT", # The deployment id for the AOAI API
```

> Need Qwen, Gemini, non‑visual GPT‑4, or even **OpenAI CUA Operator** as a AppAgent? See the [model guide](https://microsoft.github.io/UFO/supported_models/overview/).

### 📔 Step 3: Additional Setting for RAG (optional).
If you want to enhance UFO's ability with external knowledge, you can optionally configure it with an external database for retrieval augmented generation (RAG) in the `ufo/config/config.yaml` file. 

We provide the following options for RAG to enhance UFO's capabilities:
- [Offline Help Document](https://microsoft.github.io/UFO/advanced_usage/reinforce_appagent/learning_from_help_document/) Enable UFO to retrieve information from offline help documents.
- [Online Bing Search Engine](https://microsoft.github.io/UFO/advanced_usage/reinforce_appagent/learning_from_bing_search/): Enhance UFO's capabilities by utilizing the most up-to-date online search results.
- [Self-Experience](https://microsoft.github.io/UFO/advanced_usage/reinforce_appagent/experience_learning/): Save task completion trajectories into UFO's memory for future reference.
- [User-Demonstration](https://microsoft.github.io/UFO/advanced_usage/reinforce_appagent/learning_from_demonstration/): Boost UFO's capabilities through user demonstration.

Consult their respective documentation for more information on how to configure these settings.


###  Step 3 🎥: Execution Logs 

You can find the screenshots taken and request & response logs in the following folder:
```
./ufo/logs/<your_task_name>/
```
You may use them to debug, replay, or analyze the agent output.


## 2. Prepare UFO Task Files

Open `process_scripts.py` and update the configuration section at the bottom of the file to match your environment:

```python
# --- 在这里直接修改为你需要的路径 ---
master_excel_path = r"D:\code\SheetCopilot\dataset\dataset.xlsx"  # The path of dataset.xlsx in SheetCoplit 
task_sheets_dir = r"D:\code\SheetCopilot\dataset\task_sheets"     # The path of task_sheets in SheetCoplit 
output_base_dir = r"D:\code\UFO\benchmark\sheetcoplit\tasks" # The path you want to save your result
# --- 修改结束 ---
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
Step 1: Config SheetCoplit on `[SheetCoplit root dir]/agent/config/config.yaml`
```yaml
path:
  source_path: ../dataset/task_sheets
  save_path: ./DIR_to_Save
  task_path: ../dataset/dataset.xlsx
  gt_path: ../dataset/task_sheet_answers_v2
```
- Change the save path to the path of your results(such absolute path of `benchmark/sheetcoplit/tasks/agent_task_outputs`)

then cd `SheetCoplit/agent`, and run the evaluation script:

```bash
python evaluation.py --config ./config/config.yaml
```

## 5. Monitor in UFO
When ufo running in Remote Desktop/VM, you can config SMTP service to monitor progress.

Config under `./ufo/config/config_dev.yaml`
```yaml
# SMTP
MONITOR: False # SMTP or not
MACHINE_ID: DEV # Will be shown in Email tilte
SEND_POINT: 1,2 # An email will be sent once when the completed quantity exists in SEND_POINT. SEND_POINT is comma-separated.
FROM_EMAIL: 1921761583@qq.com # Email send
TO_EMAIL:  # Email receive
SENDER_PASSWORD:  # The email authorization code for the mailbox corresponding to the SMTP_SERVER
SMTP_SERVER: smtp.qq.com # SMTP server

```

You can find final output results under `evaluation/outputs`