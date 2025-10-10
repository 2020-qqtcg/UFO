import json
import os


def create_help_document(json_path, start_line, output_md_path):
    """
    Generates a Markdown help document from a JSON file. The JSON value is expected
    to be a list where the first item is a subtitle and the second is the description.

    Args:
        json_path (str): The path to the input JSON file.
        start_line (str): The string to be used for the main title.
        output_md_path (str): The path for the generated output Markdown file.
    """
    # Step 1: Create the title from the start_line variable
    # This removes the "Let's walk through " part to create a cleaner title.
    title = start_line.replace("Let's walk through ", "")

    # Step 2: Read the step data from the JSON file
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            steps_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: JSON file not found at '{json_path}'")
        return
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{json_path}'. Please check the file format.")
        return
    except Exception as e:
        print(f"An error occurred while reading the JSON file: {e}")
        return

    # Step 3: Build the Markdown content as a list of strings
    markdown_content = []

    # Add the title as a top-level heading (H1)
    markdown_content.append(f"# {title.capitalize()}\n")

    # --- MODIFIED SECTION START ---
    # Add each step from the JSON data
    for image_path, value_list in steps_data.items():
        # Check if the value is a list with at least two elements
        if isinstance(value_list, list) and len(value_list) >= 2:
            subtitle = value_list[0]
            description = value_list[1]

            # Add the subtitle as an H3 heading for better structure
            markdown_content.append(f"## {subtitle.capitalize()}\n")

            # Add the descriptive text for the step
            markdown_content.append(f"{description}\n")

            # It's good practice to use forward slashes for paths in Markdown
            formatted_path = image_path.replace('\\', '/')

            # Add the image, using the main description as its alt text.
            markdown_content.append(f"![{description}]({formatted_path})\n")
        else:
            # Handle cases where the JSON format might be different than expected
            print(f"Warning: Skipping item with key '{image_path}' due to unexpected format.")
    # --- MODIFIED SECTION END ---

    # Step 4: Write the collected content to the output Markdown file
    try:
        with open(output_md_path, 'w', encoding='utf-8') as f:
            # Join all parts with a newline for good spacing between steps
            f.write("\n".join(markdown_content))
        print(f"Successfully generated help document: {output_md_path}")
    except Exception as e:
        print(f"An error occurred while writing the output file: {e}")


# --- Example Usage ---
if __name__ == "__main__":
    # Define the paths to your input and output files.
    # You should change these to match your actual file locations.
    json_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\document_step.json"
    title = "Let's walk through how to add the © symbol to your Excel spreadsheet."
    output_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\help_document.md"


    # Call the main function to generate the document
    create_help_document(json_file, title, output_file)