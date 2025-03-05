# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.


from typing import Any, Dict, Optional, Tuple

from ufo.agents.agent.basic import BasicAgent
from ufo.agents.states.evaluaton_agent_state import EvaluatonAgentStatus
from ufo.config.config import Config
from ufo.prompter.tutorial_gen_prompter import TutorialGenAgentPrompter
from ufo.utils import json_parser, print_with_color, markdown_parser

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
    ) -> TutorialGenAgentPrompter:
        """
        Get the prompter for the agent.
        """

        return TutorialGenAgentPrompter(
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
        self, request: str, log_path: str, output_path: str, eva_all_screenshots: bool = True
    ) -> Tuple[Dict[str, str], float]:
        """
        Evaluate the task completion.
        :param log_path: The path to the log file.
        :return: The evaluation result and the cost of LLM.
        """

        message = self.message_constructor(
            log_path=log_path, request=request, eva_all_screenshots=eva_all_screenshots
        )
        result, cost = self.get_response(
            message=message, namescope="app", use_backup_engine=True
        )

        result = markdown_parser(result)

        with open(output_path, "w") as f:
            f.write(result)
            print_with_color(f"Successfully write tutorial content to: {output_path}", color="green")

        return result, cost

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
        main_prompt=configs["TUTORIAL_PROMPT"],
        example_prompt="",
        api_prompt=configs["API_PROMPT"],
    )

    request = "Add a column to the right of the 'Example Applications' column in the table."
    log_path = r"logs/test_gen/turn_dark_mode_on_or_off_76d8740a-8fa2-48d5-bd66-c841372a7235"
    output_path = r"logs/test_gen/turn_dark_mode_on_or_off_76d8740a-8fa2-48d5-bd66-c841372a7235/tutorial.md"
    results = gen_agent.generate(
        request=request, log_path=log_path, output_path=output_path, eva_all_screenshots=True
    )
