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
mkdir outputs
python evaluation.py --model custom --setting custom
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
SENDER_PASSWORD:  # Password or Key used to send email
SMTP_SERVER: smtp.qq.com # SMTP server

```

You can find final output results under `evaluation/outputs`