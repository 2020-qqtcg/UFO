import json
import os
import glob
import re # Import regular expressions for potential filename parsing

# --- Configuration ---
# PLEASE UPDATE THESE PATHS ACCORDING TO YOUR ENVIRONMENT
dataset_file_path = r"C:\Users\v-yuhangxie\repos\UFO\benchmark\SpreadsheetBench\data\all_data_912\all_data_912\dataset.json" # Path to your dataset.json file
input_base_dir = 'C:/Users/v-yuhangxie/repos/UFO/benchmark/SpreadsheetBench/data/all_data_912/all_data_912'               # Base directory where spreadsheet folders (like 'spreadsheet/59196') reside
output_dir = 'C:/Users/v-yuhangxie/repos/UFO/benchmark/SpreadsheetBench/data/all_data_912/outputs/custom_custom'        # Directory where the final output .xlsx files will be saved

# --- End Configuration ---


PROMPT_FORMAT_SINGLE = """You need to solve the given spreadsheet manipulation question, which contains six types of information:
- instruction: The question about spreadsheet manipulation.
- spreadsheet_path: The path of the spreadsheet file you need to manipulate.
- spreadsheet_content: The first few rows of the content of speadsheet file.
- instruction_type: There are two values (Cell-Level Manipulation, Sheet-Level Manipulation) used to indicate whether the answer to this question applies only to specific cells or to the entire worksheet.
- answer_position: The position need to be modified or filled. For Cell-Level Manipulation questions, this field is filled with the cell position; for Sheet-Level Manipulation, it is the maximum range of cells you need to modify. You only need to modify or fill in values within the cell range specified by answer_position.
- output_path: You need to generate the modified spreadsheet file in this new path.

Below is the spreadsheet manipulation question you need to solve:
### instruction
{instruction}

### instruction_type
{instruction_type}

### answer_position
answer_position: {answer_position}

answer_position specifies the location where you fill in the answer. Multiple ranges are separated by commas (','). Within each range, if there is an exclamation mark ('!'), the part to the left of the '!' indicates the sheet, and the part to the right indicates the range.

Do not save the file after execution.
"""

def process_dataset(dataset_path, base_input_dir, output_base_dir, tasks_output_dir,start_n,end_n):
    """
    Reads the dataset JSON, finds corresponding Excel files, and generates
    individual JSON files based on the specified rules.

    Args:
        dataset_path (str): Path to the dataset.json file.
        base_input_dir (str): Base directory containing the spreadsheet data folders.
        output_base_dir (str): Directory to save the generated JSON files.
        tasks_output_dir (str): Directory to save tasks output files.
    """
    # Ensure the output directory exists
    os.makedirs(output_base_dir, exist_ok=True)
    print(f"Output directory '{output_base_dir}' ensured.")

    try:
        with open(dataset_path, 'r', encoding='utf-8') as f:
            dataset = json.load(f)
    except FileNotFoundError:
        print(f"Error: Dataset file not found at '{dataset_path}'")
        return
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{dataset_path}'. Check its format.")
        return
    except Exception as e:
        print(f"An unexpected error occurred while reading '{dataset_path}': {e}")
        return

    if not isinstance(dataset, list):
        # Handle case where the top level might not be a list (adjust if needed)
        if isinstance(dataset, dict) and 'id' in dataset: # Check if it's a single entry object
             dataset = [dataset]
        else:
             print(f"Error: Expected a list of entries in '{dataset_path}', but got {type(dataset)}.")
             return


    processed_count = 0
    skipped_entries = 0

    # Iterate through each entry in the dataset
    for entry in dataset[start_n-1:end_n]:
        # print(entry)
        try:
            # print(entry_id)
            entry_id = entry.get('id')
            instruction = entry.get('instruction')
            instruction_type = entry.get('instruction_type')
            answer_position = entry.get('answer_position')
            spreadsheet_path_suffix = entry.get('spreadsheet_path') # e.g., "spreadsheet/59196"

            if not all([entry_id, instruction, spreadsheet_path_suffix]):
                print(f"Warning: Skipping entry due to missing 'id', 'instruction', or 'spreadsheet_path'. Entry: {entry}")
                skipped_entries += 1
                continue

            # Construct the full path to the directory containing the excel files for this entry
            # Use os.path.normpath to handle potential mixed slashes (e.g., "dataflow/outputs/datasUFO\files")
            data_dir = os.path.normpath(os.path.join(base_input_dir, spreadsheet_path_suffix))

            # Define the pattern to find the input excel files (m_n_input.xlsx)
            # Use entry_id in the pattern. '*' matches the 'm' prefix (1, 2, 3, ...)
            excel_pattern = os.path.join(data_dir, f"*_{entry_id}_input.xlsx")

            # Find all matching excel files
            found_excel_files = glob.glob(excel_pattern)

            if not found_excel_files:
                print(f"Warning: No input Excel files found for entry ID '{entry_id}' in directory '{data_dir}' using pattern '{excel_pattern}'.")
                continue

            # Process each found excel file
            for input_excel_path in found_excel_files:
                # Extract the base filename (e.g., "1_59196_input.xlsx")
                base_filename = os.path.basename(input_excel_path)

                # --- Determine the output JSON filename ---
                # Replace '_input.xlsx' with '_input.json'
                output_json_filename = base_filename.replace('_input.xlsx', '_input.json')
                output_json_path = os.path.join(tasks_output_dir, output_json_filename)

                # --- Determine the 'save_as' path ---
                # Create a corresponding output excel filename (e.g., "1_59196_output.xlsx")
                # Place it in the *output* directory for clarity, unless specified otherwise
                output_excel_filename = base_filename.replace('_input.xlsx', '_output.xlsx')
                save_as_path = os.path.normpath(os.path.join(output_base_dir, output_excel_filename))
                # If you strictly need the 'D:\example.xlsx' format from the example, uncomment the line below:
                # save_as_path = r"D:\example.xlsx" # Use raw string for Windows paths

                task = PROMPT_FORMAT_SINGLE.format_map({
                    'instruction': instruction,
                    'instruction_type': instruction_type,
                    'answer_position': answer_position
                })

                # --- Create the output JSON data structure ---
                output_data = {
                    "task": task,
                    # Use normalized path for object field for consistency
                    "object": os.path.normpath(input_excel_path),
                    "close": "True", # Default value as string
                    "save_as": save_as_path
                }

                # --- Write the output JSON file ---
                try:
                    with open(output_json_path, 'w', encoding='utf-8') as outfile:
                        json.dump(output_data, outfile, indent=4, ensure_ascii=False)
                    print(f"  Successfully generated: '{output_json_path}'")
                    processed_count += 1
                except IOError as e:
                    print(f"  Error: Could not write JSON file '{output_json_path}'. Reason: {e}")
                except Exception as e:
                    print(f"  An unexpected error occurred while writing '{output_json_path}': {e}")

        except KeyError as e:
            print(f"Warning: Skipping entry due to missing key: {e}. Entry: {entry}")
            skipped_entries += 1
        except Exception as e:
            print(f"An unexpected error occurred while processing entry: {entry}. Error: {e}")
            skipped_entries += 1


    print("\n--- Processing Summary ---")
    print(f"Total entries in dataset: {len(dataset)}")
    print(f"Generated JSON files: {processed_count}")
    print(f"Skipped entries (due to errors or missing data): {skipped_entries}")
    print("--------------------------")


# --- Main execution ---
if __name__ == "__main__":
    # start_n从1开始，从start_n~end_n的元素
    start_n=1
    end_n=100
    tasks_dir = f"C:/Users/v-yuhangxie/repos/UFO/benchmark/SpreadsheetBench/tasks_912_{start_n}-{end_n}"
    # 如果该路径不存在，则创建
    os.makedirs(tasks_dir, exist_ok=True)
    print("Starting dataset processing...")
    process_dataset(dataset_file_path, input_base_dir, output_dir, tasks_dir,start_n,end_n)
    print("Processing finished.")