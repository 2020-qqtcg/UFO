import os
import json
import base64
import time
from typing import List, Dict, Any
from ufo.llm.openai_utils import send_request_ufo, send_request_ufo_o3
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


def process_and_evaluate_steps_object(root_directory, model_name,schema):
    """
    Traverses a directory, processes document_step.json files,
    and sends the data for evaluation.
    """
    # Define the JSON schema for the expected response from the model
    with open(schema, 'r') as file:
        schema_judge = json.load(file)

    # Use os.walk to traverse all directories and files
    for root, dirs, files in os.walk(root_directory):
        # We are looking for the specific file 'document_step.json'
        if 'document_step.json' in files:
            # Check if the parent directory is named 'document'
            if os.path.basename(root) == 'document_cost':
                json_path = os.path.join(root, 'document_step.json')
                request_path=os.path.join(root, 'request.json')
                # request_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\document\request.json"
                with open(request_path, 'r', encoding='utf-8') as f:
                    request_dict = json.load(f)
                    request=request_dict["request"]

                # json_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\document\document_step.json"
                print(f"\nProcessing file: {json_path}")

                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        steps_data = json.load(f)
                        print(steps_data)
                except (json.JSONDecodeError, FileNotFoundError) as e:
                    print(f"Error reading or parsing {json_path}: {e}")
                    continue

                # Get all steps and exclude the first and last ones
                # .items() preserves insertion order in Python 3.7+
                middle_steps = list(steps_data.items())[1:-1]

                if not middle_steps:
                    print("No middle steps to process. Skipping.")
                    continue

                task_overview=f'''## Task Overview
You are an evaluation system. You are provided with an Excel task request and the details for each step from the corresponding instructional document.
Your task is to evaluate the quality of this instructional document. You need to rate each of the six criteria on a scale from 1 to 5 and provide a clear reason for each rating. The definitions and scoring guidelines for the six criteria are as follows:

<1> Clarity
Each instructional document step is clearly and explicitly described, with no ambiguity. 
"Clear" means that based on the step's textual description and screenshot, it is easy to understand how to perform the operation. "Unclear" or ambiguous means that even with the textual description and screenshot, it is still uncertain how to proceed with the operation.
A step is considered clear if the user can accurately perform the action without hesitation. For example, a step like “Click on the ‘File’ tab in the top-left corner,” accompanied by a screenshot showing the ‘File’ tab, is clear. A step is unclear if the text says, “Now, just use a VLOOKUP to get the prices,” without explaining what VLOOKUP is or how to structure the formula—especially for a user who may not be familiar with it.
Scoring Guide:
1: Every step in the instructional document is unclear, making the tutorial completely incomprehensible.
2: Only a few steps are clear, making it difficult to understand the tutorial as a whole.
3: Only about half of the step descriptions are clear, which hinders understanding of the entire process.
4: Only a few steps are slightly unclear, but they do not affect the overall comprehension of the process.
5: Every step in the instructional document is very clearly described and completely unambiguous.

<2> Conciseness
There are no repeated or unnecessary steps in the instructional document.
"Repeated or unnecessary steps" refer to redundant actions in the instructional document that can be removed without affecting the completion of the task.
For example: (1) if the task is to save a file, showing how to open the File menu and click “Save As” is necessary, but showing how to adjust font size before that is irrelevant.
(2) Unnecessary repetition, such as clicking the same button twice without purpose, also affects conciseness.
Scoring Guide:
1: Almost all steps in the document are irrelevant or repetitive, making it hard for viewers to focus on the task.
2: More than half of the steps are redundant or clearly irrelevant, significantly reducing the efficiency of information delivery.
3: The document contains multiple redundant or unnecessary steps that may affect the viewing experience.
4: The document is generally concise, with only a few steps that do not hinder understanding.
5: There are no redundant, repetitive, or unnecessary steps at all.

<3> Correctness
There are no incorrect steps in the instructional document. 
An incorrect step is one that is necessary for completing the task but is performed incompletely or incorrectly. 
For example: (1) selecting only part of column A when the whole column should be selected; (2) The task requires entering a formula, but the step includes an incorrect formula.
Scoring Guide:
1: Most of the necessary steps in the document are incorrect, making it impossible to complete the task.
2: Several necessary steps in the document are incorrect, significantly hindering task completion.
3: Some necessary steps in the document are incorrect, affecting the ability to complete the task.
4: The document is mostly correct, with only a few errors in necessary steps that do not affect overall task completion.
5: All necessary steps are entirely correct, with no errors or inaccuracies.

<4> Completeness
The instructional document steps are complete, with nothing missing or skipped.
"Complete" means that the instructional document includes all the necessary steps required to transition an Excel task from an incomplete state to a completed state.
For example: if the task is to remove duplicates from a worksheet, the steps must include selecting the relevant data, going to the “Data” tab, clicking “Remove Duplicates,” and confirming the settings.
If a step like selecting the data is missing, the document is incomplete.
Scording Guide:
1: The document covers almost none of the essential operations, with major steps missing.
2: The document is missing many steps, with most key operations not shown.
3: About half of the steps are missing; the viewer cannot understand the overall task.
4: The document is mostly complete, with a few minor missing steps that do not significantly affect task completion.
5: No steps are missing; all necessary actions to complete the task are covered.

<5> Sequential Order
The steps in the instructional document are presented in a clear and logical sequence, without any disordered or confusing arrangements.
"A clear and logical sequence" means each step is presented in the correct sequential order.
For example: In a task that involves filtering data, the correct order would be:(1) Select the data range,(2) Click the “Data” tab,(3) Click “Filter.”
If the document shows applying the filter before selecting the data, the order is incorrect and can lead to confusion or failure to complete the task.
Scoring Guide:
1: The steps are completely out of order, making it impossible to understand the workflow.
2: Several key steps are in the wrong order, resulting in a confusing flow.
3: Some steps are misordered, which creates moderate difficulty in understanding the process.
4: Steps are mostly in a reasonable order, with only minor deviations that don’t hinder understanding.
5: All steps are presented in a logical and coherent sequence.

<6> Text-Image Mapping
The text and images are well-aligned, with each image accurately matching its corresponding textual description.
Good alignment means the image directly supports what the text describes. For example, if the text says “Click the ‘Insert’ tab,” the screenshot should show the cursor pointing to the ‘Insert’ tab.
If the text specifies that a particular formula should be entered, but the screenshot does not show the actual formula, then the text and image are considered mismatched.
Scoring Guide:
1: The text and images are almost entirely misaligned, causing severe comprehension issues.
2: More than half of the images do not match the text, significantly hindering understanding.
3: Several mismatches exist, making it difficult to follow some steps by correlating text and images.
4: Overall alignment is good, with only minor mismatches that do not affect comprehension.
5: All text and images are perfectly aligned, creating a smooth and coherent visual-textual experience.

## Output Format
Output the results in JSON format. For each criterion, record both the score and a specific explanation for the score in the corresponding field using the following format:
{{
  "clarity": {{
    "score": ...,
    "reason": "..."
  }},
  "conciseness": {{
    "score": ...,
    "reason": "..."
  }},
  "correctness": {{
    "score": ...,
    "reason": "..."
  }},
  "completeness": {{
    "score": ...,
    "reason": "..."
  }},
  "sequential_order": {{
    "score": ...,
    "reason": "..."
  }},
  "text_image_mapping": {{
    "score": ...,
    "reason": "..."
  }}
}}


##input
Here is the request: {request}
Here are the titles, descriptions, and screenshots for each step in the document:
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
                    step_title, step_explanation = description
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
                        result_str = send_request_ufo_o3(model_name, message, schema=schema_judge)
                        time.sleep(30)
                        break
                    except Exception as e:
                        print_with_color(f"Error: {e}", "red")
                        print_with_color("Retrying...", "yellow")
                        try_count -= 1
                        time.sleep(30)
                        continue

                # Print the formatted result
                try:
                    result_json = json.loads(result_str)
                    output_result_file = os.path.join(root, "document_judge_result_object_0721.json")
                    with open(output_result_file, "w", encoding="utf-8") as f:
                        json.dump(result_json, f, ensure_ascii=False, indent=2)
                    print(result_json)

                except json.JSONDecodeError:
                    print(f"Could not parse model response: {result_str}")


def process_and_evaluate_steps_subject(root_directory, model_name,schema):
    """
    Traverses a directory, processes document_step.json files,
    and sends the data for evaluation.
    """
    # Define the JSON schema for the expected response from the model
    with open(schema, 'r') as file:
        schema_judge = json.load(file)

    # Use os.walk to traverse all directories and files
    for root, dirs, files in os.walk(root_directory):
        # We are looking for the specific file 'document_step.json'
        if 'document_step.json' in files:
            # Check if the parent directory is named 'document'
            if os.path.basename(root) == 'document_cost':
                json_path = os.path.join(root, 'document_step.json')
                request_path=os.path.join(root, 'request.json')
                # request_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\document\request.json"
                with open(request_path, 'r', encoding='utf-8') as f:
                    request_dict = json.load(f)
                    request=request_dict["request"]

                # json_path=r"C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\document\document_step.json"
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

                task_overview=f'''## Task Overview
You are an evaluation system. You are provided with an Excel task request and the details for each step from the corresponding instructional document.
Your task is to evaluate the quality of this instructional document. You need to rate each of the five criteria on a scale from 1 to 5 and provide a clear reason for each rating. The definitions and scoring guidelines for the five criteria are as follows:

<1> Understand
It was easy to understand the instructions.
Scoring Guide:
1: The instructions are extremely confusing, full of jargon or ambiguity, and nearly impossible to understand without external help.
2: Many parts of the instructions are unclear, causing me to pause frequently, feel confused, and struggle to follow.
3: The instructions are generally understandable, but some parts require repeated viewing or careful thought to comprehend.
4: The instructions are clear, with only a few minor points needing slight extra attention but not affecting overall understanding.
5: Every step in the instructions is extremely clear, intuitive, without any ambiguity, and can be understood effortlessly.

<2> Speed
It was faster to use tutorials like this to complete a task.
Scoring Guide:
1: Using this tutorial was even slower than figuring it out myself or consulting other materials; it was a complete waste of time.
2: Using this tutorial did not save any time and may have even taken longer due to misleading information.
3: The time to complete the task was about the same as figuring it out on my own; I did not feel any significant improvement in efficiency.
4: The speed of completing the task was noticeably improved, saving me a considerable amount of time.
5: Extremely efficient; compared to any other method, this tutorial saved me a massive amount of time and effort.

<3> Complete task
I was able to complete the tasks without problems.
Scoring Guide:
1: I was completely unable to complete the task based on the tutorial and was even led to an incorrect result.
2: I struggled significantly while completing the task, or I could only complete some of a part of the steps.
3: I eventually completed the task, but I encountered several problems or errors along the way that I had to solve myself.
4: I completed the task smoothly, encountering only a few minor hiccups that were easily resolved.
5: I completed all task steps flawlessly and without any obstacles.

<4> Satisfaction
I was satisfied with tutorials like this.
Scoring Guide:
1: I am very dissatisfied with this experience; the entire process was frustrating and unhelpful.
2: I am dissatisfied with this experience; I encountered more problems than I received help.
3: The overall experience was average, neither good nor bad, with no particular feelings.
4: I am satisfied with this experience; it was a positive and helpful process.
5: I am very satisfied with this experience; the process was pleasant, efficient, and far exceeded my expectations.

<5> Preference
When finding help online, I would prefer to use tutorials like this.
Scoring Guide:
1: I will actively avoid this type of tutorial in the future.
2: I would prefer to use other forms of tutorials (e.g., text-only documents, traditional videos).
3: I have no particular preference; I would use it if I come across it but would not actively seek it out.
4: In most situations, I would choose this type of tutorial first.
5: This is my most preferred format for tutorials; I will look for this type first when seeking help in the future.

## Output Format
Output the results in JSON format. For each criterion, record both the score and a specific explanation for the score in the corresponding field using the following format:
{{
  "understand": {{
    "score": ...,
    "reason": "..."
  }},
  "speed": {{
    "score": ...,
    "reason": "..."
  }},
  "complete_task": {{
    "score": ...,
    "reason": "..."
  }},
  "satisfaction": {{
    "score": ...,
    "reason": "..."
  }},
  "preference": {{
    "score": ...,
    "reason": "..."
  }}
}}

##input
Here is the request: {request}
Here are the titles, descriptions, and screenshots for each step in the document:
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
                    step_title, step_explanation = description
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
                        result_str = send_request_ufo_o3(model_name, message, schema=schema_judge)
                        time.sleep(30)
                        break
                    except Exception as e:
                        print_with_color(f"Error: {e}", "red")
                        print_with_color("Retrying...", "yellow")
                        try_count -= 1
                        time.sleep(30)
                        continue

                # Print the formatted result
                try:
                    result_json = json.loads(result_str)
                    output_result_file = os.path.join(root, "document_judge_result_subject_0721.json")
                    with open(output_result_file, "w", encoding="utf-8") as f:
                        json.dump(result_json, f, ensure_ascii=False, indent=2)
                    print(result_json)

                except json.JSONDecodeError:
                    print(f"Could not parse model response: {result_str}")


if __name__ == '__main__':
    # IMPORTANT: Set this to the root folder you want to scan.
    root_folder = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_o3\excel_complete_double"

    # Specify the model you want to use
    judge_model_name = 'dev-o3-2025-04-16'
    # judge_model_name ="dev-gpt-45-preview"# Or any other available model
    schema_object="./data/steps_schema_questionnaire_score_object.json"
    schema_subject = "./data/steps_schema_questionnaire_score_subject.json"

    if not os.path.isdir(root_folder):
        print(f"Error: Root directory not found at '{root_folder}'")
    else:
        process_and_evaluate_steps_object(root_folder, judge_model_name, schema_object)
        process_and_evaluate_steps_subject(root_folder, judge_model_name, schema_subject)
    # message = [
    #     {"role": "user", "content": "Who are you?"},
    # ]
    # model_name = 'dev-gpt-41-longco-2025-04-14'
    # i=20
    # while i>0:
    #     i=i-1
    #     try:
    #         responses = send_request_ufo(
    #             model_name, message
    #         )
    #         print(responses)
    #     except Exception as e:
    #         print_with_color(f"Error: {e}", "red")
    #         time.sleep(30)
