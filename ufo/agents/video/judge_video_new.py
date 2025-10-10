import os
import json
import base64
import time
from typing import List, Dict, Any
from ufo.llm.openai_utils import send_request_ufo
from colorama import Fore, Style, init


def print_with_color(text: str, color: str = "", end: str = "\n") -> None:
    """
    Print text with specified color using ANSI escape codes from Colorama library.

    :param text: The text to print.
    :param color: The color of the text (options: red, green, yellow, blue, magenta, cyan, white, black).
    """
    color_mapping = {
        "red": Fore.RED,
        "green": Fore.GREEN,
        "yellow": Fore.YELLOW,
        "blue": Fore.BLUE,
        "magenta": Fore.MAGENTA,
        "cyan": Fore.CYAN,
        "white": Fore.WHITE,
        "black": Fore.BLACK,
    }

    selected_color = color_mapping.get(color.lower(), "")
    colored_text = selected_color + text + Style.RESET_ALL

    print(colored_text, end=end)


# --- Helper and Main Logic ---

def image_to_base64(image_path: str) -> str:
    """Encodes an image file to a Base64 string."""
    try:
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            return encoded_string
    except FileNotFoundError:
        print(f"Warning: Image file not found at {image_path}")
        return None


def process_and_evaluate_steps_video_new(root_directory, model_name, schema):
    """
    Traverses a directory, processes video_step.json files,
    and sends the data for evaluation.
    """
    # Define the JSON schema for the expected response from the model
    with open(schema, 'r') as file:
        schema_judge = json.load(file)

    # Use os.walk to traverse all directories and files
    for root, dirs, files in os.walk(root_directory):
        # We are looking for the specific file 'video_step.json'
        if 'video_step.json' in files:
            # Check if the parent directory is named 'video'
            if os.path.basename(root) == 'video_cost':

                # Check if the result file already exists
                output_result_file_path = os.path.join(root, "video_judge_result_new.json")
                if os.path.exists(output_result_file_path):
                    print_with_color(f"Skipping because result file already exists: {output_result_file_path}",
                                     "yellow")
                    continue

                json_path = os.path.join(root, 'video_step.json')
                request_path = os.path.join(root, 'request.json')
                # request_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\document\request.json"
                with open(request_path, 'r', encoding='utf-8') as f:
                    request_dict = json.load(f)
                    request = request_dict["request"]

                # json_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\video_step.json"
                print(f"\nProcessing file: {json_path}")

                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        steps_data = json.load(f)
                        # print(steps_data)
                except (json.JSONDecodeError, FileNotFoundError) as e:
                    print(f"Error reading or parsing {json_path}: {e}")
                    continue

                # Get all steps and exclude the first and last ones
                # .items() preserves insertion order in Python 3.7+
                middle_steps = list(steps_data.items())[1:-1]

                if not middle_steps:
                    print("No middle steps to process. Skipping.")
                    continue

                task_overview = f'''## Task Overview
You are an evaluation system. You are provided with an Excel task request and the details for each step from the corresponding instructional video.
Your task is to evaluate the quality of this instructional video. You need to rate each of the seven criteria on a scale from 1 to 5 and provide a clear reason for each rating. The definitions and scoring guidelines for the six criteria are as follows:

<1> Usability
I found the tutorial easy to understand and could follow the steps without confusion.
Scoring Guide:
1: The tutorial was confusing, full of jargon or unclear explanations, and the steps were hard to understand or perform, even after multiple attempts.
2: Some information in the tutorial was unclear, and I had to guess what to do in several steps, making it hard to complete the process smoothly.
3: The tutorial was generally understandable, but I had to pause often, rewatch, or experiment repeatedly to grasp the content and steps.
4. The tutorial was clear, with logically ordered steps, and I could understand and perform them with little extra effort.
5: Every step of the tutorial was extremely clear, intuitive, and unambiguous. I could follow and complete the process effortlessly.

<2> Correctness
There are no incorrect steps in the instructional video. An incorrect step is one that is necessary for completing the task but is performed incompletely or incorrectly. For example: (1) selecting only part of column A when the whole column should be selected; (2) The task requires entering a formula, but the step includes an incorrect formula.
Scoring Guide:
1: Most of the necessary steps in the video are incorrect, making it impossible to complete the task.
2: Several necessary steps in the video are incorrect, significantly hindering task completion.
3: Some necessary steps in the video are incorrect, affecting the ability to complete the task.
4: The video is mostly correct, with only a few errors in necessary steps that do not affect overall task completion.
5: All necessary steps are entirely correct, with no errors or inaccuracies.

<3> Interactivity
The video provided sufficient cues at key steps (e.g., mouse pointers, highlighted boxes to indicate the operation area), which helped user notice important actions in time.
Scoring Guide:
1: There were no cues at all, making many key operations easy to miss.
2: The cues were insufficient, which caused me to miss or misunderstand some important steps.
3: There were some cues, but they were not timely or noticeable enough.
4. The video gave clear cues at key points, which helped me complete the operations.
5: The cues were very timely and clear, effectively guiding me through every key step.

<4> Design Quality
Was the video well-structured and accessible — with a clear beginning, middle, and end, logical step-by-step organization, and friendly presentation elements (e.g., subtitles, narration, color scheme, and pacing) suitable for users?
Scoring Guide:
1: The video was confusing, with poor pacing, disorganized structure, no clear beginning, middle, or end, and unhelpful subtitles or visuals that made it hard to understand.
2: Some design elements (e.g., subtitles, narration, colors) or the structural flow were poorly arranged, and the steps lacked logical order, which made comprehension difficult for me.
3: The structure and design were mostly acceptable, and the subtitles or narration were somewhat helpful, but I still had to put in effort to understand and follow.
4. The video was clearly structured, with logically ordered steps and a clear beginning, middle, and end. The subtitles, pacing, and visual design supported my understanding.
5: The video was excellently designed, with a coherent and logical structure from beginning to end, and user-centered design elements like subtitles, narration, colors, and pacing that made it extremely easy to understand and follow.

<5> Transferability
I was able to apply the methods learned in the video to similar situations.
Scoring Guide:
1: The tutorial knowledge was too specific and not applicable to other tasks.
2: Only a small part of the knowledge could be transferred, and it required extra effort.
3: The knowledge was somewhat transferable but required my own adaptation.
4: I could apply the tutorial content directly to similar contexts.
5: I could easily transfer the knowledge to other similar tasks.

<6> Task Completion and Satisfaction
I was able to complete the task smoothly using the tutorial and was satisfied with the overall experience.
Scoring Guide:
1: I was completely unable to complete the task based on the tutorial; the entire process was frustrating and unhelpful, and therefore I am very dissatisfied with this experience.
2. I struggled significantly while completing the task, encountering more problems than I received help, and therefore I am dissatisfied with this experience.
3. I eventually completed the task, but had to solve some problems myself along the way, so the overall experience was average, neither good nor bad.
4. I completed the task smoothly, the process was positive and helpful, and therefore I am satisfied with this experience.
5: I completed all tasks flawlessly and without any obstacles; the process was pleasant and efficient, which makes me very satisfied with this experience.

<7> Efficiency and Preference
I find this type of tutorial to be efficient, and therefore I would prefer to use it in the future.
Scoring Guide:
1: I believe using this tutorial is a complete waste of time, even slower than figuring it out myself, and therefore I will actively avoid it in the future.
2: I believe this tutorial did not save time and was not efficient, therefore I would prefer to use other forms of tutorials.
3: I believe the efficiency improvement is not noticeable, about the same as figuring it out myself, therefore I have no particular preference and would not actively seek out this type of tutorial.
4. I believe this tutorial significantly improved my efficiency, and therefore in most situations, I would choose it first.
5: I believe this tutorial is extremely efficient and saved me a lot of time; it has become my most preferred format, and I will look for it first in the future.

## Output Format
Output the results in JSON format. For each criterion, record both the score and a specific explanation for the score in the corresponding field using the following format:
{{
  "usability": {{
    "score": ...,
    "reason": "..."
  }},
  "correctness": {{
    "score": ...,
    "reason": "..."
  }},
  "interactivity": {{
    "score": ...,
    "reason": "..."
  }},
  "design_quality": {{
    "score": ...,
    "reason": "..."
  }},
  "transferability": {{
    "score": ...,
    "reason": "..."
  }},
  "completion_satisfaction": {{
    "score": ...,
    "reason": "..."
  }},
  "efficiency_preference": {{
    "score": ...,
    "reason": "..."
  }}
}}

##input
Here is the request: {request}
Here are the titles, descriptions, and screenshots for each step in the video:
  '''
                # Construct the message payload for the API
                user_content = [
                    {
                        "type": "text",
                        "text": task_overview
                    }
                ]

                for image_path, description in middle_steps:
                    # Add the text description for the step
                    step_title = description["title"]
                    step_explanation = description["voiceover_script"]

                    user_content.append({
                        "type": "text",
                        "text": f"Step title: {step_title}\nStep description: {step_explanation}"
                    })

                    # Encode and add the image for the step
                    base64_image = image_to_base64(image_path)
                    if base64_image:
                        user_content.append({
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        })

                message = [
                    {"role": "user", "content": user_content}
                ]
                try_count = 20
                while try_count > 0:
                    try:
                        # Send the request and get the result
                        result_str = send_request_ufo(model_name, message, schema=schema_judge)
                        time.sleep(5)
                        break
                    except Exception as e:
                        print_with_color(f"Error: {e}", "red")
                        print_with_color("Retrying...", "yellow")
                        try_count -= 1
                        time.sleep(5)
                        continue

                # Print the formatted result
                try:
                    result_json = json.loads(result_str)
                    output_result_file = os.path.join(root, "video_judge_result_new.json")
                    with open(output_result_file, "w", encoding="utf-8") as f:
                        json.dump(result_json, f, ensure_ascii=False, indent=2)
                    print(result_json)

                except json.JSONDecodeError:
                    print(f"Could not parse model response: {result_str}")


if __name__ == '__main__':
    # 1. Define a list containing the paths to the three root folders.
    # bing_search_query_410002497
    # IMPORTANT: Set these to the root folders you want to scan.
    # root_folders = [
    #     r"C:\Users\v-yuhangxie\OneDrive - Microsoft\log_result\2025_0712_qabench_completed_new",
    #     r"C:\Users\v-yuhangxie\OneDrive - Microsoft\log_result\20250716_m365_complete",
    #     r"C:\Users\v-yuhangxie\OneDrive - Microsoft\log_result\20250716_bing_search_completed"
    # ]

    # root_folders = [
    #     # r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\qabench_completion_double"
    #     r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\bing_completion_double"
    # ]

    root_folders = [
        r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_qabench_4.1_cost_complete_double",
        r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_bing_4.1_cost_complete_double"
    ]

    # 2. Specify the model you want to use (this part is unchanged).
    judge_model_name = 'dev-gpt-41-longco-2025-04-14'
    # judge_model_name ="dev-gpt-45-preview"
    schema_new = "./data/steps_schema_questionnaire_score_video.json"

    # 3. Loop through the list of folders and process each one.
    for folder_path in root_folders:
        print(f"\n--- Processing Folder: {folder_path} ---")
        if not os.path.isdir(folder_path):
            print(f"Error: Root directory not found at '{folder_path}'")
            continue  # Skip to the next folder if this one doesn't exist

        # Call the processing function for the current folder in the loop.
        process_and_evaluate_steps_video_new(folder_path, judge_model_name, schema_new)

    print("\n--- All folders have been processed. ---")