# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
import time
import shutil
import json
import os
from typing import Any, Dict, Optional, Tuple

from ufo.agents.agent.basic import BasicAgent
from ufo.agents.states.evaluaton_agent_state import EvaluatonAgentStatus
from ufo.config.config import Config
from ufo.prompter.demo_gen_prompter import DemoGenAgentPrompter
from ufo.utils import print_with_color
from ufo.agents.gen_tutorial.tool.get_request import extract_and_clean_requests
from ufo.agents.gen_tutorial.tool.gen_document_md import create_help_document
from ufo.agents.gen_tutorial.tool.gen_document_html import create_html_document
from ufo.agents.gen_tutorial.tool.gen_document_html_base64_pic import create_html_document_base64

configs = Config.get_instance().config_data


class TutorialGenAgent(BasicAgent):
    """
    The agent for evaluation.
    """

    def __init__(
            self,
            name: str,
            app_root_name: str,
            is_visual: bool,
            main_prompt: str,
            example_prompt: str,
            api_prompt: str,
    ):
        """
        Initialize the FollowAgent.
        :agent_type: The type of the agent.
        :is_visual: The flag indicating whether the agent is visual or not.
        """

        super().__init__(name=name)

        self._app_root_name = app_root_name
        self.prompter = self.get_prompter(
            is_visual,
            main_prompt,
            example_prompt,
            api_prompt,
            app_root_name,
        )

    def get_prompter(
            self,
            is_visual,
            prompt_template: str,
            example_prompt_template: str,
            api_prompt_template: str,
            root_name: Optional[str] = None,
    ) -> DemoGenAgentPrompter:
        """
        Get the prompter for the agent.
        """

        return DemoGenAgentPrompter(
            is_visual=is_visual,
            prompt_template=prompt_template,
            example_prompt_template=example_prompt_template,
            api_prompt_template=api_prompt_template,
            root_name=root_name,
        )

    def message_constructor(
            self, log_path: str, request: str, eva_all_screenshots: bool = True
    ) -> Dict[str, Any]:
        """
        Construct the message.
        :param log_path: The path to the log file.
        :param request: The request.
        :param eva_all_screenshots: The flag indicating whether to evaluate all screenshots.
        :return: The message.
        """

        agent_prompt_system_message = self.prompter.system_prompt_construction()

        agent_prompt_user_message = self.prompter.user_content_construction(
            log_path=log_path, request=request, eva_all_screenshots=eva_all_screenshots
        )

        agent_prompt_message = self.prompter.prompt_construction(
            agent_prompt_system_message, agent_prompt_user_message
        )

        return agent_prompt_message

    @property
    def status_manager(self) -> EvaluatonAgentStatus:
        """
        Get the status manager.
        """

        return EvaluatonAgentStatus

    def generate(
            self, request: str, log_path: str, output_path: str, eva_all_screenshots: bool = True, schema: dict = None
    ) -> Tuple[Dict[str, str], float]:
        """
        Evaluate the task completion.
        :param log_path: The path to the log file.
        :return: The evaluation result and the cost of LLM.
        """

        message = self.message_constructor(
            log_path=log_path, request=request, eva_all_screenshots=eva_all_screenshots
        )
        result,prompt_tokens,completion_tokens,cost,time_taken_seconds = self.get_response_schema_cost(
            message=message, schema=schema, namescope="app", use_backup_engine=True
        )

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result)
            print_with_color(f"Successfully write tutorial content to: {output_path}", color="green")

        return result,prompt_tokens,completion_tokens,cost,time_taken_seconds

    def process_comfirmation(self) -> None:
        """
        Comfirmation, currently do nothing.
        """
        pass


# The following code is used for testing the agent.
if __name__ == "__main__":

    gen_agent = TutorialGenAgent(
        name="tutorial_gen_agent",
        app_root_name="WINWORD.EXE",
        is_visual=True,
        main_prompt=configs["DEMO_PROMPT_DOCUMENT"],
        example_prompt="",
        api_prompt=configs["API_PROMPT"],
    )


    base_path = r"C:\Users\v-yuhangxie\UFO_1011\logs\20251011_try_complete_double"
    for folder_name in os.listdir(base_path):
        log_path = os.path.join(base_path, folder_name)

        # 确保当前项目是一个文件夹
        if os.path.isdir(log_path):
            print(f"✅ 正在访问文件夹: {log_path}")


        md_file_path = os.path.join(log_path, "output.md")
        if not os.path.exists(md_file_path):
            print(f"{md_file_path} 不存在，跳过")
            continue

        request = extract_and_clean_requests(md_file_path)

        # 创建 output_folder 路径：log_path 下的 "document"
        output_folder = os.path.join(log_path, "document")
        os.makedirs(output_folder, exist_ok=True)  # 如果不存在就创建

        # 生成两个输出文件的完整路径
        step_output_path = os.path.join(output_folder, "document_demo_step.json")
        step_cost_path = os.path.join(output_folder, "document_demo_cost.json")

        request_output_path = os.path.join(output_folder, "request.json")
        with open(request_output_path, "w", encoding="utf-8") as f:
            request_dict={"request":request}
            json.dump(request_dict, f, ensure_ascii=False, indent=2)

        with open('./ufo/agents/gen_tutorial/data/steps_schema_document.json', 'r') as file:
            schema = json.load(file)

        result,prompt_tokens,completion_tokens,cost,time_taken_seconds = gen_agent.generate(
            request=request, log_path=log_path, output_path=step_output_path,eva_all_screenshots=True, schema=schema
        )
        cost_dict = {"llm_request":{"prompt_tokens":prompt_tokens,"completion_tokens":completion_tokens,"cost":cost,"time_taken_seconds":time_taken_seconds}}
        print("-----------------------------------")

        with open(step_output_path, 'r', encoding='utf-8') as f:
            image_step_dict = json.load(f)

        time_start = time.time()
        task_title = image_step_dict["task_title"]
        task_title = "Help Document: " + task_title

        # 构建包含完整源图片路径的字典，用于HTML生成

        image_step_path_dict_document = {}
        path = os.path.join(log_path, f"action_step1.png")
        text = ["input file", "Let's see how to do this together!"]
        image_step_path_dict_document[path] = text
        i=1
        for key, value in image_step_dict["steps"].items():
            path = os.path.join(log_path, f"action_step{key}_selected_controls_mouse.png")
            written_explanation = value['written_explanation']
            title = value['title']
            image_step_path_dict_document[path] = [f"step{i}: {title}", f"{written_explanation}"]
            i=i+1
        path = os.path.join(log_path, f"action_step_final.png")
        image_step_path_dict_document[path] = ["output file", "You've now successfully completed the task!"]

        document_json_path = os.path.join(output_folder, "document_step.json")
        # 保存为 JSON 文件 (用于HTML)
        with open(document_json_path, "w", encoding="utf-8") as f:
            json.dump(image_step_path_dict_document, f, ensure_ascii=False, indent=2)

        # --- 新增: 处理MD文件和图片 ---
        # 1. 创建md输出文件夹
        md_output_folder = os.path.join(output_folder, "md")
        os.makedirs(md_output_folder, exist_ok=True)

        # 2. 复制图片并创建新的字典 (用于MD)
        image_step_path_dict_for_md = {}
        for src_path, text_data in image_step_path_dict_document.items():
            if os.path.exists(src_path):
                img_filename = os.path.basename(src_path)
                dest_path = os.path.join(md_output_folder, img_filename)

                # 复制图片到md文件夹
                shutil.copy2(src_path, dest_path)

                # 在新字典中使用相对路径 (文件名)
                image_step_path_dict_for_md[img_filename] = text_data
            else:
                print_with_color(f"Warning: Source image not found, skipping: {src_path}", "yellow")

        # 3. 将新的MD数据字典保存为JSON
        md_json_path = os.path.join(md_output_folder, "document_step_for_md.json")
        with open(md_json_path, "w", encoding="utf-8") as f:
            json.dump(image_step_path_dict_for_md, f, ensure_ascii=False, indent=2)

        # 4. 指定MD文件的最终输出路径
        output_file_document_md = os.path.join(md_output_folder, "help_document.md")

        # --- 文件生成 ---
        # 使用新的md_json_path为MD文件生成提供数据
        create_help_document(md_json_path, task_title, output_file_document_md)
        print_with_color(f"Successfully created Markdown document at: {output_file_document_md}", "green")

        # HTML文件生成逻辑保持不变
        output_file_document_html = os.path.join(output_folder, "help_document.html")
        output_file_document_html_base64 = os.path.join(output_folder, "help_document_base64.html")
        create_html_document(document_json_path, task_title, output_file_document_html)
        create_html_document_base64(document_json_path, task_title, output_file_document_html_base64)

        time_end = time.time()
        time_taken_seconds = time_end - time_start
        cost_dict["gen_document_time"]=time_taken_seconds
        with open(step_cost_path, "w", encoding="utf-8") as f:
            json.dump(cost_dict, f, ensure_ascii=False, indent=2)
