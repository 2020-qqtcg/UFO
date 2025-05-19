from utils import compare_workbooks
import tqdm
import yaml
import pandas as pd
import os
import time
# import numpy as np #不再需要numpy，因为动作计数相关的统计移除了
from datetime import datetime
import argparse
from collections import defaultdict

USE_NO_AND_SHEETNAME = False


def evaluate(config):
    task_path_config = config['path']['task_path']  # Renamed to avoid conflict with loop variable
    gt_path = config['path']['gt_path']
    save_path = config['path']['save_path']
    eval_result_path = os.path.join(config['path']['save_path'], 'eval_result.yaml')

    print("Evaluate the results at ", save_path)
    if os.path.exists(eval_result_path):
        with open(eval_result_path, 'r') as f:
            eval_result = yaml.load(f, Loader=yaml.Loader)
    else:
        eval_result = {"check_result_each_repeat": {}}

    task_df = pd.read_excel(task_path_config, header=0)

    print(
        "\033[0;36;40m========================================================\nEvaluate task result: {}\033[0m\n".format(
            save_path))

    for repeat_id in range(1, config["repeat"] + 1):
        t = time.time()
        if eval_result["check_result_each_repeat"].get(repeat_id, None) is None:
            eval_result["check_result_each_repeat"][repeat_id] = {
                "matched_gt_lst": [],
                "checked_list": [],
                "exec_success_list": [],  # 任务的输出文件存在即视为执行成功
                "success_list": [],  # 任务的输出文件与GT匹配即视为成功
                "checked_list_by_cate": defaultdict(list),
                "exec_success_list_by_cate": defaultdict(list),
                "success_list_by_cate": defaultdict(list),
                "error_log": [],
                "eval_results": {}
            }

        check_result = eval_result["check_result_each_repeat"][repeat_id]

        # 确定 save_path 下有多少个任务目录
        if os.path.exists(save_path) and os.path.isdir(save_path):
            num_tasks_in_save_path = len(
                [x for x in os.listdir(save_path) if os.path.isdir(os.path.join(save_path, x))])
        else:
            print(
                f"\033[0;31;40mWarning: save_path '{save_path}' does not exist or is not a directory. No tasks will be evaluated from there.\033[0m")
            num_tasks_in_save_path = 0

        # 评估的任务数量以 task_df 中的任务为准
        num_tasks_to_evaluate = len(task_df)

        remaining_task_cnt = num_tasks_to_evaluate - len(check_result["checked_list"])
        # assert remaining_task_cnt >= 0, "Task counting mismatch or already evaluated all tasks." # Allow re-evaluation or partial evaluation

        if remaining_task_cnt <= 0 and num_tasks_to_evaluate > 0:
            print(
                f"\033[0;33;40mAll {num_tasks_to_evaluate} tasks already evaluated for repeat {repeat_id}. Skipping.\033[0m")
            # Still print summary if available
            if check_result["eval_results"]:
                print("\033[0;33;40mSummary for Repeat {}:\033[0m".format(repeat_id))
                for k, v_ in check_result["eval_results"].items():
                    print("{}: {}".format(k, v_))
                print("========================================================\n")
            continue  # Skip to next repeat if no tasks left for this one
        elif num_tasks_to_evaluate == 0:
            print(f"\033[0;31;40mNo tasks found in task_df from '{task_path_config}'. Cannot evaluate.\033[0m")
            continue

        with tqdm.tqdm(total=remaining_task_cnt,
                       desc=f"Processing remaining {remaining_task_cnt}/{num_tasks_to_evaluate} results of repeat {repeat_id}") as pbar:
            for index, row in task_df.iloc[:].iterrows():
                task_identifier_for_checked_list = f"{index + 1}_{row['Sheet Name']}" if not USE_NO_AND_SHEETNAME else f"{row['No.']}_{row['Sheet Name']}"

                if task_identifier_for_checked_list in check_result["checked_list"]:  # Use consistent identifier
                    continue

                # Result file name based on task_df
                if USE_NO_AND_SHEETNAME:
                    task_name = f"{row['No.']}_{row['Sheet Name']}"
                else:
                    task_name = f"{index + 1}_{row['Sheet Name']}"

                # Path to the folder containing results for this specific task
                current_task_save_folder = os.path.join(save_path, task_name)

                # Path to the agent's output Excel file for this task and repeat
                res_path = os.path.join(current_task_save_folder, f"{task_name}_{repeat_id}.xlsx")

                # Check if the result xlsx file exists
                res_file_exists = os.path.exists(res_path)
                cates = str(row['Categories']).split(', ')  # Ensure categories is string

                if not os.path.exists(current_task_save_folder):
                    # print(f"Task folder {current_task_save_folder} not found. Skipping task.")
                    # Still mark as checked to prevent re-processing in this run.
                    # The actual output file non-existence will handle exec_success and success.
                    pass  # Will be handled by res_file_exists check below.

                if res_file_exists:
                    check_result["exec_success_list"].append(task_identifier_for_checked_list)
                    for cate in cates:
                        check_result["exec_success_list_by_cate"][cate].append(task_identifier_for_checked_list)

                if res_file_exists:
                    # Compare the result with all reference solutions.
                    # All reference solutions for one sheet is placed under a folder with the same name.

                    # Load GTs
                    # Ensure 'No.' column exists and is used correctly, or handle if missing
                    gt_task_folder_name = f"{row.get('No.', index + 1)}_{row['Sheet Name']}"  # Fallback to index if 'No.' is missing
                    gt_folder_this_task = os.path.join(gt_path, row['Sheet Name'], gt_task_folder_name)

                    if not os.path.exists(gt_folder_this_task):
                        check_result["error_log"].append(
                            f"GT folder {gt_folder_this_task} not exists for task {task_identifier_for_checked_list}")
                        # continue # Decide if you want to skip or just log error and mark checked
                    else:
                        gt_files_found = False
                        for gt_file_name in [x for x in os.listdir(gt_folder_this_task) if
                                             x.endswith('.xlsx') and "$" not in x]:
                            gt_files_found = True
                            gt_file_path = os.path.join(gt_folder_this_task, gt_file_name)
                            check_board_path = os.path.join(gt_folder_this_task,
                                                            gt_file_name.replace(".xlsx", "_check.yaml"))

                            if not os.path.exists(check_board_path):
                                check_result["error_log"].append(
                                    f"Check_board {check_board_path} not exists for GT {gt_file_name}")
                                continue  # Skip this GT if its check_board is missing

                            with open(check_board_path, 'r') as f:
                                check_board = yaml.load(f, Loader=yaml.Loader)

                            if not os.path.exists(gt_file_path):
                                # This case should be rare if os.listdir found it, but good practice
                                check_result["error_log"].append(f"GT file {gt_file_path} listed but not accessible")
                                continue

                            """
                            Comparing.......
                            """
                            try:
                                check_res = compare_workbooks(gt_file_path, res_path, check_board["check_board"])
                                """
                                Comparing.....................
                                """
                                # If checking is successful
                                if check_res[1]:  # check_res[1] is True for a match
                                    check_result["success_list"].append(task_identifier_for_checked_list)
                                    for cate in cates:
                                        check_result["success_list_by_cate"][cate].append(
                                            task_identifier_for_checked_list)
                                    check_result["matched_gt_lst"].append(gt_file_name)
                                    break  # Matched GT found. Stop checking other GTs for this task
                            except Exception as e:
                                check_result["error_log"].append(
                                    f"Error comparing {os.path.basename(res_path)} with {gt_file_name}: {str(e)}")

                        if not gt_files_found and os.path.exists(gt_folder_this_task):
                            check_result["error_log"].append(
                                f"No GT .xlsx files found in {gt_folder_this_task} for task {task_identifier_for_checked_list}")
                else:  # res_file does not exist
                    check_result["error_log"].append(
                        f"Result file {res_path} not found for task {task_identifier_for_checked_list}")

                check_result["checked_list"].append(task_identifier_for_checked_list)  # Use consistent identifier
                for cate in cates:
                    check_result["checked_list_by_cate"][cate].append(task_identifier_for_checked_list)

                # Save intermediate results
                with open(eval_result_path, 'w') as f:
                    yaml.dump(eval_result, f, allow_unicode=True,
                              sort_keys=False)  # Added sort_keys=False for consistency

                pbar.update(1)

        print("\033[0;33;40mEvaluation for Repeat {} has finished. Time elapse: {:.2f}s\033[0m".format(repeat_id,
                                                                                                       time.time() - t))
        if check_result["error_log"]:
            print("Error Log:\n{}\n".format('\n'.join(x for x in check_result["error_log"])))

        exec_success_cnt = len(set(check_result["exec_success_list"]))  # Use set to count unique tasks
        success_cnt = len(set(check_result["success_list"]))  # Use set to count unique tasks
        total_checked_unique = len(set(check_result["checked_list"]))  # Use set for unique tasks

        check_result["eval_results"]["Total_Unique_Checked"] = total_checked_unique

        if total_checked_unique > 0:
            check_result["eval_results"]["Exec@1 (File Exists)"] = exec_success_cnt / total_checked_unique
            check_result["eval_results"]["Pass@1 (Matched GT)"] = success_cnt / total_checked_unique
        else:
            check_result["eval_results"]["Exec@1 (File Exists)"] = 0
            check_result["eval_results"]["Pass@1 (Matched GT)"] = 0

        for k_cate, v_list in check_result["checked_list_by_cate"].items():
            cate_total_unique = len(set(v_list))
            if cate_total_unique > 0:
                cate_exec_success_cnt = len(set(check_result["exec_success_list_by_cate"].get(k_cate, [])))
                cate_pass_cnt = len(set(check_result["success_list_by_cate"].get(k_cate, [])))
                check_result["eval_results"][
                    f"{k_cate} Exec & Pass (Unique)"] = "{:d}/{:d} ({:.2%}) & {:d}/{:d} ({:.2%})".format(
                    cate_exec_success_cnt, cate_total_unique,
                    cate_exec_success_cnt / cate_total_unique if cate_total_unique else 0,
                    cate_pass_cnt, cate_total_unique, cate_pass_cnt / cate_total_unique if cate_total_unique else 0
                )
            else:
                check_result["eval_results"][f"{k_cate} Exec & Pass (Unique)"] = "0/0 (0.00%) & 0/0 (0.00%)"

        # Storing lists as strings for YAML readability (original behavior)
        # Using unique task identifiers from sets before joining.
        check_result["matched_gt_lst"] = ', '.join(str(x) for x in sorted(list(set(check_result["matched_gt_lst"]))))
        check_result["checked_list"] = ', '.join(str(x) for x in sorted(list(set(check_result["checked_list"]))))
        check_result["success_list"] = ', '.join(str(x) for x in sorted(list(set(check_result["success_list"]))))
        check_result["exec_success_list"] = ', '.join(
            str(x) for x in sorted(list(set(check_result["exec_success_list"]))))

        for k_cate in check_result['checked_list_by_cate']:
            check_result['checked_list_by_cate'][k_cate] = ', '.join(
                str(x) for x in sorted(list(set(check_result['checked_list_by_cate'][k_cate]))))
        for k_cate in check_result['success_list_by_cate']:
            check_result['success_list_by_cate'][k_cate] = ', '.join(
                str(x) for x in sorted(list(set(check_result['success_list_by_cate'][k_cate]))))
        for k_cate in check_result['exec_success_list_by_cate']:
            check_result['exec_success_list_by_cate'][k_cate] = ', '.join(
                str(x) for x in sorted(list(set(check_result['exec_success_list_by_cate'][k_cate]))))

        # Action statistics are removed

        for k_res, v_res in check_result["eval_results"].items():
            print("{}: {}".format(k_res, v_res))

        print("========================================================\n")

        # Save the metrics to the eval_result and save it
        with open(eval_result_path, 'w') as f:
            yaml.dump(eval_result, f, allow_unicode=True, sort_keys=False)

    print("{} have been evaluated on {}... . Time: {}".format(save_path, gt_path, datetime.now().strftime("%H:%M:%S")))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process config.')
    parser.add_argument('--config', '-c', default="./config/config.yaml", type=str, help='path to config file')
    args = parser.parse_args()

    if not os.path.exists(args.config):
        print(f"Error: Config file '{args.config}' not found.")
    else:
        with open(args.config, 'r') as f:
            config = yaml.load(f, Loader=yaml.Loader)

        # Ensure save_path exists for eval_result.yaml
        if 'path' in config and 'save_path' in config['path']:
            os.makedirs(config['path']['save_path'], exist_ok=True)
            evaluate(config)
            print("Evaluate {}".format(config["path"]["save_path"]))
        else:
            print("Error: 'path' or 'save_path' not defined in config.")