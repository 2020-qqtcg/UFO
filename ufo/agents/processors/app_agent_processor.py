# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import json
import os
from dataclasses import asdict, dataclass, field
from typing import TYPE_CHECKING, Any, ClassVar, Dict, List, Optional, Union

from pywinauto.controls.uiawrapper import UIAWrapper

from ufo import utils
from ufo.agents.processors.actions import ActionSequence, BaseControlLog, OneStepAction
from ufo.agents.processors.basic import BaseProcessor
from ufo.automator.ui_control import ui_tree
from ufo.automator.ui_control.control_filter import ControlFilterFactory
from ufo.automator.ui_control.grounding.basic import BasicGrounding
from ufo.config.config import Config
from ufo.module.context import Context, ContextNames

if TYPE_CHECKING:
    from ufo.agents.agent.app_agent import AppAgent

configs = Config.get_instance().config_data

if configs is not None:
    CONTROL_BACKEND = configs.get("CONTROL_BACKEND", ["uia"])
    BACKEND = "win32" if "win32" in CONTROL_BACKEND else "uia"


@dataclass
class AppAgentAdditionalMemory:
    """
    The additional memory data for the AppAgent.
    """

    Step: int
    RoundStep: int
    AgentStep: int
    Round: int
    Subtask: str
    SubtaskIndex: int
    FunctionCall: List[str]
    Action: List[Dict[str, Any]]
    ActionSuccess: List[Dict[str, Any]]
    ActionType: List[str]
    Request: str
    Agent: str
    AgentName: str
    Application: str
    Cost: float
    Results: str
    error: str
    time_cost: Dict[str, float]
    ControlLog: Dict[str, Any]
    UserConfirm: Optional[str] = None


@dataclass
class AppAgentControlLog(BaseControlLog):
    """
    The control log data for the AppAgent.
    """

    control_friendly_class_name: str = ""
    control_coordinates: Dict[str, int] = field(default_factory=dict)


@dataclass
class ControlInfoRecorder:
    """
    The control meta information recorder for the current application window.
    """

    recording_fields: ClassVar[List[str]] = [
        "control_text",
        "control_type" if BACKEND == "uia" else "control_class",
        "control_rect",
        "source",
    ]

    application_windows_info: Dict[str, Any] = field(default_factory=dict)
    uia_controls_info: List[Dict[str, Any]] = field(default_factory=list)
    grounding_controls_info: List[Dict[str, Any]] = field(default_factory=list)
    merged_controls_info: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class AppAgentRequestLog:
    """
    The request log data for the AppAgent.
    """

    step: int
    dynamic_examples: List[str]
    experience_examples: List[str]
    demonstration_examples: List[str]
    offline_docs: str
    online_docs: str
    dynamic_knowledge: str
    image_list: List[str]
    prev_subtask: List[str]
    plan: List[str]
    request: str
    control_info: List[Dict[str, str]]
    subtask: str
    current_application: str
    host_message: str
    blackboard_prompt: List[str]
    last_success_actions: List[Dict[str, Any]]
    include_last_screenshot: bool
    prompt: Dict[str, Any]
    control_info_recording: Dict[str, Any]


class AppAgentProcessor(BaseProcessor):
    """
    The processor for the app agent at a single step.
    """

    # Add a class-level cache for Excel layout detection
    _excel_layout_cache = {}
    _cache_timestamp = 0

    def __init__(
        self,
        agent: "AppAgent",
        context: Context,
        ground_service=Optional[BasicGrounding],
    ) -> None:
        """
        Initialize the app agent processor.
        :param agent: The app agent who executes the processor.
        :param context: The context of the session.
        """

        super().__init__(agent=agent, context=context)

        self.app_agent = agent
        self.host_agent = agent.host

        self._annotation_dict = None
        self._control_info = None
        self._operation = ""
        self._args = {}
        self._image_url = []
        self.control_filter_factory = ControlFilterFactory()
        self.control_recorder = ControlInfoRecorder()
        self.filtered_annotation_dict = None
        self.screenshot_save_path = None
        self.grounding_service = ground_service

    def print_step_info(self) -> None:
        """
        Print the step information.
        """
        utils.print_with_color(
            "Round {round_num}, Step {step}, AppAgent: Completing the subtask [{subtask}] on application [{application}].".format(
                round_num=self.round_num + 1,
                step=self.round_step + 1,
                subtask=self.subtask,
                application=self.application_process_name,
            ),
            "magenta",
        )

    def get_control_list(self, screenshot_path: str) -> List[UIAWrapper]:
        """
        Get the control list from the annotation dictionary.
        :param screenshot_path: The path to the clean screenshot.
        :return: The list of control items.
        """

        api_backend = None
        grounding_backend = None

        control_detection_backend = configs.get("CONTROL_BACKEND", ["uia"])

        if "uia" in control_detection_backend:
            api_backend = "uia"
        elif "win32" in control_detection_backend:
            api_backend = "win32"

        if "omniparser" in control_detection_backend:
            grounding_backend = "omniparser"

        if api_backend is not None:
            api_control_list = (
                self.control_inspector.find_control_elements_in_descendants(
                    self.application_window,
                    control_type_list=configs.get("CONTROL_LIST", []),
                    class_name_list=configs.get("CONTROL_LIST", []),
                )
            )
        else:
            api_control_list = []

        api_control_dict = {
            i + 1: control for i, control in enumerate(api_control_list)
        }

        # print(control_detection_backend, grounding_backend, screenshot_path)

        if grounding_backend == "omniparser" and self.grounding_service is not None:
            self.grounding_service: BasicGrounding

            onmiparser_configs = configs.get("OMNIPARSER", {})

            # print(onmiparser_configs)

            grounding_control_list = (
                self.grounding_service.convert_to_virtual_uia_elements(
                    image_path=screenshot_path,
                    application_window=self.application_window,
                    box_threshold=onmiparser_configs.get("BOX_THRESHOLD", 0.05),
                    iou_threshold=onmiparser_configs.get("IOU_THRESHOLD", 0.1),
                    use_paddleocr=onmiparser_configs.get("USE_PADDLEOCR", True),
                    imgsz=onmiparser_configs.get("IMGSZ", 640),
                )
            )
        else:
            grounding_control_list = []

        grounding_control_dict = {
            i + 1: control for i, control in enumerate(grounding_control_list)
        }

        merged_control_list = self.photographer.merge_control_list(
            api_control_list,
            grounding_control_list,
            iou_overlap_threshold=configs.get("IOU_THRESHOLD_FOR_MERGE", 0.1),
        )

        merged_control_dict = {
            i + 1: control for i, control in enumerate(merged_control_list)
        }

        # Record the control information for the uia controls.
        self.control_recorder.uia_controls_info = (
            self.control_inspector.get_control_info_list_of_dict(
                api_control_dict, ControlInfoRecorder.recording_fields
            )
        )

        # Record the control information for the grounding controls.
        self.control_recorder.grounding_controls_info = (
            self.control_inspector.get_control_info_list_of_dict(
                grounding_control_dict, ControlInfoRecorder.recording_fields
            )
        )

        # Record the control information for the merged controls.
        self.control_recorder.merged_controls_info = (
            self.control_inspector.get_control_info_list_of_dict(
                merged_control_dict, ControlInfoRecorder.recording_fields
            )
        )

        return merged_control_list

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def capture_screenshot(self) -> None:
        """
        Capture the screenshot.
        """

        # Define the paths for the screenshots saved.
        screenshot_save_path = self.log_path + f"action_step{self.session_step}.png"
        self.screenshot_save_path = screenshot_save_path

        annotated_screenshot_save_path = (
            self.log_path + f"action_step{self.session_step}_annotated.png"
        )
        concat_screenshot_save_path = (
            self.log_path + f"action_step{self.session_step}_concat.png"
        )

        self._memory_data.add_values_from_dict(
            {
                "CleanScreenshot": screenshot_save_path,
                "AnnotatedScreenshot": annotated_screenshot_save_path,
                "ConcatScreenshot": concat_screenshot_save_path,
            }
        )

        self.photographer.capture_app_window_screenshot(
            self.application_window, save_path=screenshot_save_path
        )

        # Record the control information for the current application window.
        self.control_recorder.application_windows_info = (
            self.control_inspector.get_control_info(
                self.application_window, field_list=ControlInfoRecorder.recording_fields
            )
        )

        # # Get the control elements in the application window if the control items are not provided for reannotation.
        # if type(self.control_reannotate) == list and len(self.control_reannotate) > 0:
        #     control_list = self.control_reannotate
        # else:
        #     control_list = self.control_inspector.find_control_elements_in_descendants(
        #         self.application_window,
        #         control_type_list=configs.get("CONTROL_LIST", []),
        #         class_name_list=configs.get("CONTROL_LIST", []),
        #     )

        control_list = self.get_control_list(screenshot_save_path)

        # Get the annotation dictionary for the control items, in a format of {control_label: control_element}.
        self._annotation_dict = self.photographer.get_annotation_dict(
            self.application_window, control_list, annotation_type="number"
        )

        # Get API annotation dictionary
        original_control_list = configs.get("CONTROL_LIST", [])
        configs["CONTROL_LIST"] = ["Button", "Edit", "TabItem", "Document", "ListItem", "MenuItem", "ScrollBar", "TreeItem", "Hyperlink", "ComboBox", "RadioButton", "Spinner", "CheckBox", "DataItem"]
        api_control_list = self.get_control_list(screenshot_save_path)
        configs["CONTROL_LIST"] = original_control_list
        self._api_annotation_dict = self.photographer.get_annotation_dict(
            self.application_window, api_control_list, annotation_type="number"
        )

        # Attempt to filter out irrelevant control items based on the previous plan.
        self.filtered_annotation_dict = self.get_filtered_annotation_dict(
            self._annotation_dict
        )

        # Capture the screenshot of the selected control items with annotation and save it.
        self.photographer.capture_app_window_screenshot_with_annotation_dict(
            self.application_window,
            self.filtered_annotation_dict,
            annotation_type="number",
            save_path=annotated_screenshot_save_path,
        )

        if configs.get("SAVE_UI_TREE", False):
            if self.application_window is not None:
                step_ui_tree = ui_tree.UITree(self.application_window)
                step_ui_tree.save_ui_tree_to_json(
                    os.path.join(
                        self.ui_tree_path, f"ui_tree_step{self.session_step}.json"
                    )
                )

        if configs.get("SAVE_FULL_SCREEN", False):

            desktop_save_path = (
                self.log_path + f"desktop_action_step{self.session_step}.png"
            )

            self._memory_data.add_values_from_dict(
                {"DesktopCleanScreenshot": desktop_save_path}
            )

            # Capture the desktop screenshot for all screens.
            self.photographer.capture_desktop_screen_screenshot(
                all_screens=True, save_path=desktop_save_path
            )

        # If the configuration is set to include the last screenshot with selected controls tagged, save the last screenshot.
        if configs.get("INCLUDE_LAST_SCREENSHOT", True):
            last_screenshot_save_path = (
                self.log_path + f"action_step{self.session_step - 1}.png"
            )
            last_control_screenshot_save_path = (
                self.log_path
                + f"action_step{self.session_step - 1}_selected_controls.png"
            )
            self._image_url += [
                self.photographer.encode_image_from_path(
                    last_control_screenshot_save_path
                    if os.path.exists(last_control_screenshot_save_path)
                    else last_screenshot_save_path
                )
            ]

        # Whether to concatenate the screenshots of clean screenshot and annotated screenshot into one image.
        if configs.get("CONCAT_SCREENSHOT", False):
            self.photographer.concat_screenshots(
                screenshot_save_path,
                annotated_screenshot_save_path,
                concat_screenshot_save_path,
            )
            self._image_url += [
                self.photographer.encode_image_from_path(concat_screenshot_save_path)
            ]
        else:
            screenshot_url = self.photographer.encode_image_from_path(
                screenshot_save_path
            )
            screenshot_annotated_url = self.photographer.encode_image_from_path(
                annotated_screenshot_save_path
            )
            self._image_url += [screenshot_url, screenshot_annotated_url]

        # Save the XML file for the current state.
        if configs.get("LOG_XML", False):
            self._save_to_xml()

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def get_control_info(self) -> None:
        """
        Get the control information.
        """

        # Get the control information for the control items and the filtered control items, in a format of list of dictionaries.
        self._control_info = self.control_inspector.get_control_info_list_of_dict(
            self._annotation_dict,
            ["control_text", "control_type" if BACKEND == "uia" else "control_class"],
        )
        self.filtered_control_info = (
            self.control_inspector.get_control_info_list_of_dict(
                self.filtered_annotation_dict,
                [
                    "control_text",
                    "control_type" if BACKEND == "uia" else "control_class",
                ],
            )
        )

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def get_prompt_message(self) -> None:
        """
        Get the prompt message for the AppAgent.
        """

        experience_results, demonstration_results = (
            self.app_agent.demonstration_prompt_helper(request=self.subtask)
        )

        retrieved_results = experience_results + demonstration_results

        # Get the external knowledge prompt for the AppAgent using the offline and online retrievers.

        offline_docs, online_docs = self.app_agent.external_knowledge_prompt_helper(
            self.subtask,
            configs.get("RAG_OFFLINE_DOCS_RETRIEVED_TOPK", 0),
            configs.get("RAG_ONLINE_RETRIEVED_TOPK", 0),
        )

        # print(offline_docs, online_docs)

        external_knowledge_prompt = offline_docs + online_docs

        if not self.app_agent.blackboard.is_empty():
            blackboard_prompt = self.app_agent.blackboard.blackboard_to_prompt()
        else:
            blackboard_prompt = []

        # Get the last successful actions of the AppAgent.
        last_success_actions = self.get_last_success_actions()

        action_keys = ["Function", "Args", "ControlText", "Results", "RepeatTimes"]

        filtered_last_success_actions = [
            {key: action.get(key, "") for key in action_keys}
            for action in last_success_actions
        ]

        # Construct the prompt message for the AppAgent.
        self._prompt_message = self.app_agent.message_constructor(
            dynamic_examples=retrieved_results,
            dynamic_knowledge=external_knowledge_prompt,
            image_list=self._image_url,
            control_info=self.filtered_control_info,
            prev_subtask=self.previous_subtasks,
            plan=self.prev_plan,
            request=self.request,
            subtask=self.subtask,
            current_application=self.application_process_name,
            host_message=self.host_message,
            blackboard_prompt=blackboard_prompt,
            last_success_actions=filtered_last_success_actions,
            include_last_screenshot=configs.get("INCLUDE_LAST_SCREENSHOT", True),
        )

        # Log the prompt message. Only save them in debug mode.
        request_data = AppAgentRequestLog(
            step=self.session_step,
            experience_examples=experience_results,
            demonstration_examples=demonstration_results,
            dynamic_examples=retrieved_results,
            offline_docs=offline_docs,
            online_docs=online_docs,
            dynamic_knowledge=external_knowledge_prompt,
            image_list=self._image_url,
            prev_subtask=self.previous_subtasks,
            plan=self.prev_plan,
            request=self.request,
            control_info=self.filtered_control_info,
            subtask=self.subtask,
            current_application=self.application_process_name,
            host_message=self.host_message,
            blackboard_prompt=blackboard_prompt,
            last_success_actions=filtered_last_success_actions,
            include_last_screenshot=configs.get("INCLUDE_LAST_SCREENSHOT", True),
            prompt=self._prompt_message,
            control_info_recording=asdict(self.control_recorder),
        )

        request_log_str = json.dumps(asdict(request_data), ensure_ascii=False)
        self.request_logger.debug(request_log_str)

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def get_response(self) -> None:
        """
        Get the response from the LLM.
        """

        retry = 0
        while retry < configs.get("JSON_PARSING_RETRY", 3):
            # Try to get the response from the LLM. If an error occurs, catch the exception and log the error.
            self._response, self.cost = self.app_agent.get_response(
                self._prompt_message, "APPAGENT", use_backup_engine=True
            )

            try:
                self.app_agent.response_to_dict(self._response)
                break
            except Exception as e:
                print("Error in parsing response: ", e)
                retry += 1

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def parse_response(self) -> None:
        """
        Parse the response.
        """

        self._response_json = self.app_agent.response_to_dict(self._response)

        self.control_label = self._response_json.get("ControlLabel", "")
        self.control_text = self._response_json.get("ControlText", "")
        self._operation = self._response_json.get("Function", "")
        self.question_list = self._response_json.get("Questions", [])
        self._args = utils.revise_line_breaks(self._response_json.get("Args", ""))

        # Convert the plan from a string to a list if the plan is a string.
        self.plan = self.string2list(self._response_json.get("Plan", ""))
        self._response_json["Plan"] = self.plan

        self.status = self._response_json.get("Status", "")
        self.app_agent.print_response(
            response_dict=self._response_json, print_action=True
        )

    @BaseProcessor.exception_capture
    @BaseProcessor.method_timer
    def execute_action(self) -> None:
        """
        Execute the action.
        """

        action = OneStepAction(
            function=self._operation,
            args=self._args,
            control_label=self._control_label,
            control_text=self.control_text,
            after_status=self.status,
        )
        control_selected = self._annotation_dict.get(self._control_label, None)

        # Save the screenshot of the tagged selected control.
        self.capture_control_screenshot(control_selected, action)

        self.actions: ActionSequence = ActionSequence(actions=[action])
        self.actions.execute_all(
            puppeteer=self.app_agent.Puppeteer,
            control_dict=self._annotation_dict,
            application_window=self.application_window,
        )

        if self.is_application_closed():
            utils.print_with_color("Warning: The application is closed.", "yellow")
            self.status = "FINISH"

    def capture_control_screenshot(
        self, control_selected: Union[UIAWrapper, List[UIAWrapper]], action: OneStepAction = None
    ) -> None:
        """
        Capture the screenshot of the selected control.
        :param control_selected: The selected control item or a list of selected control items.
        :param action: The action being executed (for API calls).
        """
        control_screenshot_save_path = (
            self.log_path + f"action_step{self.session_step}_selected_controls.png"
        )

        self._memory_data.add_values_from_dict(
            {"SelectedControlScreenshot": control_screenshot_save_path}
        )

        # Check if this is an API call without control selection
        if control_selected is None and action is not None:
            # Try to get API affected coordinates before action execution
            api_coords = self._get_api_affected_coordinates(action)
            if api_coords:
                # Use the new method for API coordinate-based screenshot
                self.photographer.capture_app_window_screenshot_with_rectangle_from_adjusted_coords(
                    self.application_window,
                    control_adjusted_coords=api_coords,
                    save_path=control_screenshot_save_path,
                    background_screenshot_path=self.screenshot_save_path,
                )
                return

        # Original logic for UI control selection
        sub_control_list = (
            control_selected
            if isinstance(control_selected, list)
            else [control_selected] if control_selected is not None
            else []
        )

        if sub_control_list:
            self.photographer.capture_app_window_screenshot_with_rectangle(
                self.application_window,
                sub_control_list=sub_control_list,
                save_path=control_screenshot_save_path,
                background_screenshot_path=self.screenshot_save_path,
            )
        else:
            # If no control and no API coordinates, just copy the clean screenshot
            import shutil
            if self.screenshot_save_path and os.path.exists(self.screenshot_save_path):
                shutil.copy2(self.screenshot_save_path, control_screenshot_save_path)

    def _get_api_affected_coordinates(self, action: OneStepAction) -> List[Dict[str, float]]:
        """
        Get the affected coordinates for API calls.
        :param action: The executed action.
        :return: List of coordinate dictionaries for drawing rectangles.
        """
        function_name = action.function
        args = action.args

        # Excel range operation APIs - use unified coordinate calculation method
        if function_name == "select_table_range":
            return self._get_excel_range_coordinates(args)
        elif function_name == "insert_excel_table":
            return self._get_excel_insert_table_coordinates(args)
        elif function_name == "get_range_values":
            return self._get_excel_range_coordinates(args)  # Reuse range method
        elif function_name == "auto_fill":
            return self._get_excel_auto_fill_coordinates(args)
        elif function_name == "set_cell_value":
            return self._get_excel_single_cell_coordinates(args)
        elif function_name == "reorder_columns":
            return self._get_excel_columns_coordinates(args)
        
        return []

    def _get_excel_insert_table_coordinates(self, args: Dict[str, Any]) -> List[Dict[str, float]]:
        """
        Calculate coordinates for area affected by insert_excel_table API
        :param args: insert_excel_table parameters (sheet_name, table, start_row, start_col)
        :return: Coordinates of inserted table area
        """
        try:
            start_row = args.get("start_row", 1)
            start_col = args.get("start_col", 1)
            table = args.get("table", [[]])
            
            if not table or not table[0]:
                return []
            
            # Calculate table end position
            end_row = start_row + len(table) - 1
            end_col = start_col + len(table[0]) - 1
            

            
            # Reuse existing range coordinate calculation method
            range_args = {
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col
            }
            return self._get_excel_range_coordinates(range_args)
            
        except Exception as e:
            print(f"❌ Error calculating insert table coordinates: {e}")
            return []

    def _get_excel_auto_fill_coordinates(self, args: Dict[str, Any]) -> List[Dict[str, float]]:
        """
        Calculate coordinates for area affected by auto_fill API
        :param args: auto_fill parameters (sheet_name, start_row, start_col, end_row, end_col)
        :return: Coordinates of auto-fill area
        """
        try:
            start_row = args.get("start_row", 1)
            start_col = args.get("start_col", 1)
            end_row = args.get("end_row", start_row)
            end_col = args.get("end_col", start_col)
            
            # Handle column letter conversion
            if isinstance(start_col, str):
                start_col = self._col_letter_to_num(start_col)
            if isinstance(end_col, str):
                end_col = self._col_letter_to_num(end_col)
            

            
            # Reuse existing range coordinate calculation method
            range_args = {
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col
            }
            return self._get_excel_range_coordinates(range_args)
            
        except Exception as e:
            print(f"❌ Error calculating auto fill coordinates: {e}")
            return []

    def _get_excel_single_cell_coordinates(self, args: Dict[str, Any]) -> List[Dict[str, float]]:
        """
        Calculate coordinates for single cell of set_cell_value API
        :param args: set_cell_value parameters (sheet_name, row, col, value, is_formula)
        :return: Coordinates of single cell
        """
        try:
            row = args.get("row", 1)
            col = args.get("col", 1)
            
            # Handle column letter conversion
            if isinstance(col, str):
                col = self._col_letter_to_num(col)
            

            
            # Reuse existing range coordinate calculation method (range of single cell)
            range_args = {
                "start_row": row,
                "start_col": col,
                "end_row": row,
                "end_col": col
            }
            return self._get_excel_range_coordinates(range_args)
            
        except Exception as e:
            print(f"❌ Error calculating single cell coordinates: {e}")
            return []

    def _get_excel_columns_coordinates(self, args: Dict[str, Any]) -> List[Dict[str, float]]:
        """
        Calculate coordinates for columns affected by reorder_columns API
        Uses the same logic as the API: find corresponding columns by checking first row cell content
        :param args: reorder_columns parameters (sheet_name, desired_order)
        :return: Coordinates of affected columns
        """
        try:
            desired_order = args.get("desired_order", [])
            
            if not desired_order:
        
                return []
            
            # Use the same logic as API: find first row cell content
            target_columns = []
            first_row_cells = []
            
            # Find all first row cells from _annotation_dict
            for control_label, control_element in (self._annotation_dict or {}).items():
                try:
                    # Check if this is a cell control
                    if self._is_cell_control(control_element, ""):
                        # Infer cell position
                        cell_position = self._infer_cell_position(control_element, "")
                        if cell_position:
                            cell_row, cell_col = cell_position
                            
                            # Only focus on first row cells (column header row)
                            if cell_row == 1:
                                control_text = getattr(control_element, 'window_text', lambda: '')()
                                
                                first_row_cells.append({
                                    "col": cell_col,
                                    "text": control_text.strip() if control_text else "",
                                    "element": control_element,
                                    "label": control_label
                                })
                                
                except Exception as e:
                    continue
            
            # Sort by column number, simulating API's column traversal logic
            first_row_cells.sort(key=lambda x: x["col"])
            
            # Get actual content of first row cells through Excel COM object
            found_columns = []
            try:
                # Try to get Excel COM object to read cell content
                excel_com = self.app_agent.Puppeteer.receiver_manager.com_receiver.com_object
                if excel_com:
                    active_sheet = excel_com.ActiveSheet
                    
                    for cell in first_row_cells:
                        cell_address = cell["text"]  # e.g.: "A1", "B1", "C1"
                        if cell_address:
                            try:
                                # Use COM object to read actual cell value
                                cell_value = active_sheet.Range(cell_address).Value
                                if cell_value and str(cell_value).strip() in desired_order:
                                    rect = cell["element"].rectangle()
                                    target_columns.append({
                                        "left": float(rect.left),
                                        "top": float(rect.top),
                                        "right": float(rect.right),
                                        "bottom": float(rect.bottom),
                                        "column_name": str(cell_value).strip(),
                                        "column_index": cell["col"],
                                        "cell_address": cell_address
                                    })
                                    found_columns.append(f"{str(cell_value).strip()}(col {cell['col']})")
                            except Exception:
                                # Skip if unable to read specific cell
                                continue
                else:
                    # Fall back to original logic if COM object is not available (use address matching)
                    for cell in first_row_cells:
                        if cell["text"] and cell["text"] in desired_order:
                            rect = cell["element"].rectangle()
                            target_columns.append({
                                "left": float(rect.left),
                                "top": float(rect.top),
                                "right": float(rect.right),
                                "bottom": float(rect.bottom),
                                "column_name": cell["text"],
                                "column_index": cell["col"]
                            })
                            found_columns.append(f"{cell['text']}(col {cell['col']})")
                            
            except Exception:
                # Fall back to original logic if COM object access fails
                for cell in first_row_cells:
                    if cell["text"] and cell["text"] in desired_order:
                        rect = cell["element"].rectangle()
                        target_columns.append({
                            "left": float(rect.left),
                            "top": float(rect.top),
                            "right": float(rect.right),
                            "bottom": float(rect.bottom),
                            "column_name": cell["text"],
                            "column_index": cell["col"]
                        })
                        found_columns.append(f"{cell['text']}(col {cell['col']})")
            
            if target_columns:
                # Calculate merged rectangle of all target columns (including entire columns, not just first row)
                # Need to extend to entire visible column area
                if len(target_columns) > 0:
                    # Find leftmost and rightmost columns
                    min_left = min(col["left"] for col in target_columns)
                    max_right = max(col["right"] for col in target_columns)
                    
                    # Get worksheet layout info to calculate entire column range
                    layout = self._detect_excel_interface_layout()
                    app_rect = self.application_window.rectangle()
                    
                    # Extend to entire column: from worksheet top to bottom of visible area
                    worksheet_top = layout['worksheet_top']
                    worksheet_bottom = min(app_rect.height() - 50, worksheet_top + 600)  # Estimate visible area
                    
                    merged_rect = {
                        "left": min_left - app_rect.left,
                        "top": float(worksheet_top), 
                        "right": max_right - app_rect.left,
                        "bottom": float(worksheet_bottom)
                    }
                    
                    return [merged_rect]
            
            # If specific columns cannot be found, use traditional method to estimate entire worksheet area
            layout = self._detect_excel_interface_layout()
            app_rect = self.application_window.rectangle()
            
            # Estimate entire visible worksheet area
            worksheet_rect = {
                "left": float(layout['worksheet_left']),
                "top": float(layout['worksheet_top']),
                "right": float(min(app_rect.width() - 50, layout['worksheet_left'] + 800)),
                "bottom": float(min(app_rect.height() - 50, layout['worksheet_top'] + 600))
            }
            
            return [worksheet_rect]
            
        except Exception as e:
            return []

    def clear_excel_layout_cache(self) -> None:
        """
        Clear the Excel layout detection cache to force fresh detection.
        Useful when Excel window is resized or layout changes.
        """
        self._excel_layout_cache.clear()
        self._cache_timestamp = 0


    def _detect_excel_interface_layout(self) -> Dict[str, int]:
        """
        Dynamically detect Excel interface layout by finding key controls and measuring their positions.
        Uses caching to avoid repeated expensive detection operations.
        :return: Dictionary with interface measurements.
        """
        import time
        
        # Check cache first (cache valid for 30 seconds)
        current_time = time.time()
        cache_key = id(self.application_window)
        
        if (cache_key in self._excel_layout_cache and 
            current_time - self._cache_timestamp < 30):
            return self._excel_layout_cache[cache_key]
        
        try:
            app_rect = self.application_window.rectangle()
            layout = {
                'worksheet_left': 48,     # Default fallback
                'worksheet_top': 201,     # Default fallback  
                'cell_width': 72,         # Default fallback
                'cell_height': 21         # Default fallback
            }
            
            # Search for Excel UI elements with more comprehensive criteria
            descendants = self.control_inspector.find_control_elements_in_descendants(
                self.application_window,
                control_type_list=["Edit", "ComboBox", "Text", "Button", "Header", "HeaderItem"],
                class_name_list=[]
            )
            
            name_box_rect = None
            formula_bar_rect = None
            row_headers = []
            column_headers = []
            ribbon_height = 0
            
            print(f"Analyzing {len(descendants)} UI elements for Excel layout detection...")
            
            for control in descendants:
                control_name = getattr(control.element_info, 'name', '').lower()
                control_id = getattr(control.element_info, 'automation_id', '')
                control_class = getattr(control.element_info, 'class_name', '')
                control_type = getattr(control.element_info, 'control_type', '')
                
                try:
                    rect = control.rectangle()
                    
                    # Detect name box (cell reference display)
                    if ('name box' in control_name or 
                        control_id == 'FormulaBarNameBox' or
                        control_id == 'NameBox' or
                        (control_type == 'ComboBox' and rect.top < app_rect.height() * 0.2)):
                        name_box_rect = rect
                        print(f"✓ Found name box: {control_name} ({control_id}), rect: {rect}")
                    
                    # Detect formula bar
                    elif ('formula bar' in control_name or 
                          'formula' in control_name or
                          control_id == 'FormulaBarEdit' or
                          control_id == 'FormulaEditBox' or
                          (control_type == 'Edit' and rect.top < app_rect.height() * 0.2 and 
                           rect.width() > app_rect.width() * 0.3)):
                        formula_bar_rect = rect
                        print(f"✓ Found formula bar: {control_name} ({control_id}), rect: {rect}")
                    
                    # Detect row headers (numbers on the left side)
                    elif (control_type == 'HeaderItem' and 
                          rect.left < app_rect.width() * 0.15 and  # Slightly wider tolerance
                          rect.width() < 80 and rect.height() < 60):  # More flexible sizing
                        # Additional validation: should contain numeric text or be at left edge
                        control_text = getattr(control, 'window_text', lambda: '')()
                        if (control_text.isdigit() or 
                            rect.left - app_rect.left < 100):  # Very close to left edge
                            row_headers.append(rect)
                            print(f"✓ Found row header: {control_text}, rect: {rect}")
                    
                    # Detect column headers (letters at the top) 
                    elif (control_type == 'HeaderItem' and
                          rect.top < app_rect.height() * 0.4 and  # More tolerance for ribbon variations
                          rect.width() < 250 and rect.height() < 60):
                        # Additional validation: should contain alphabetic text or be at top
                        control_text = getattr(control, 'window_text', lambda: '')()
                        if (control_text.isalpha() or 
                            rect.top - app_rect.top < app_rect.height() * 0.3):
                            column_headers.append(rect)
                            print(f"✓ Found column header: {control_text}, rect: {rect}")
                    
                    # Detect ribbon height for better top calculation
                    elif ('ribbon' in control_name.lower() or 
                          'tab' in control_name.lower() and rect.top < app_rect.height() * 0.25):
                        ribbon_height = max(ribbon_height, rect.bottom - app_rect.top)
                        
                except Exception as e:
                    continue  # Skip controls that can't provide rectangle info
            
            print(f"Detection summary: {len(row_headers)} row headers, {len(column_headers)} column headers")
            
            # Calculate worksheet_left dynamically
            if row_headers:
                # Find the rightmost row header to determine where worksheet starts
                max_right = max(header.right for header in row_headers)
                layout['worksheet_left'] = max(20, max_right - app_rect.left + 3)
                print(f"📏 Calculated worksheet_left from row headers: {layout['worksheet_left']}")
            elif name_box_rect:
                # Fallback: estimate from name box position
                layout['worksheet_left'] = max(40, name_box_rect.left - app_rect.left)
                print(f"📏 Calculated worksheet_left from name box: {layout['worksheet_left']}")
            
            # Calculate worksheet_top dynamically
            if column_headers:
                # Find the bottommost column header to determine where worksheet starts
                max_bottom = max(header.bottom for header in column_headers)
                layout['worksheet_top'] = max(120, max_bottom - app_rect.top + 3)
                print(f"📏 Calculated worksheet_top from column headers: {layout['worksheet_top']}")
            elif formula_bar_rect:
                # Fallback: estimate from formula bar
                layout['worksheet_top'] = max(150, formula_bar_rect.bottom - app_rect.top + 35)
                print(f"📏 Calculated worksheet_top from formula bar: {layout['worksheet_top']}")
            elif name_box_rect:
                # Further fallback: estimate from name box
                layout['worksheet_top'] = max(150, name_box_rect.bottom - app_rect.top + 55)
                print(f"📏 Calculated worksheet_top from name box: {layout['worksheet_top']}")
            elif ribbon_height > 0:
                # Use ribbon height as reference
                layout['worksheet_top'] = max(150, ribbon_height + 80)
                print(f"📏 Calculated worksheet_top from ribbon: {layout['worksheet_top']}")
            
            # Calculate cell dimensions dynamically with improved algorithm
            if len(column_headers) >= 2:
                # Calculate average column width from column headers
                sorted_headers = sorted(column_headers, key=lambda h: h.left)
                widths = []
                for i in range(len(sorted_headers) - 1):
                    width = sorted_headers[i + 1].left - sorted_headers[i].left
                    if 25 < width < 300:  # Broader reasonable column width range
                        widths.append(width)
                if widths:
                    # Use median instead of mean for more robust calculation
                    widths.sort()
                    layout['cell_width'] = int(widths[len(widths) // 2])
                    print(f"📏 Calculated cell_width from column headers (median): {layout['cell_width']}")
            
            if len(row_headers) >= 2:
                # Calculate average row height from row headers
                sorted_headers = sorted(row_headers, key=lambda h: h.top)
                heights = []
                for i in range(len(sorted_headers) - 1):
                    height = sorted_headers[i + 1].top - sorted_headers[i].top
                    if 12 < height < 80:  # Broader reasonable row height range
                        heights.append(height)
                if heights:
                    # Use median instead of mean for more robust calculation
                    heights.sort()
                    layout['cell_height'] = int(heights[len(heights) // 2])
                    print(f"📏 Calculated cell_height from row headers (median): {layout['cell_height']}")
            
            # Apply reasonable bounds to prevent extreme values
            layout['worksheet_left'] = max(15, min(layout['worksheet_left'], 250))
            layout['worksheet_top'] = max(80, min(layout['worksheet_top'], 500))
            layout['cell_width'] = max(25, min(layout['cell_width'], 400))
            layout['cell_height'] = max(12, min(layout['cell_height'], 120))
            
            # Cache the result
            self._excel_layout_cache[cache_key] = layout
            self._cache_timestamp = current_time
            
            print(f"🎯 Final Excel layout detected: {layout}")
            return layout
            
        except Exception as e:
            print(f"❌ Failed to detect Excel layout dynamically, using defaults: {e}")
            # Return conservative default values if detection fails
            default_layout = {
                'worksheet_left': 48,
                'worksheet_top': 201,
                'cell_width': 72,
                'cell_height': 21
            }
            # Cache the default layout too
            self._excel_layout_cache[cache_key] = default_layout
            self._cache_timestamp = current_time
            return default_layout

    def _get_excel_range_coordinates(self, args: Dict[str, Any]) -> List[Dict[str, float]]:
        """
        Calculate screen coordinates for Excel range selection using hybrid approach.
        Prioritize annotation_dict-based precise method with landmark detection as fallback.
        
        Working principle example:
        1. API call: select_table_range(start_row=1, start_col=1, end_row=3, end_col=2)  # A1:B3
        2. Search from _annotation_dict: 
           - Find controls 174 (A1), 175 (A2), 176 (A3), 189 (B1), 190 (B2), 191 (B3)
        3. Get real rectangle coordinates of each cell
        4. Calculate merged rectangle: minimum bounding box of all found cells
        5. Return precise range coordinates for screenshot marking
        
        :param args: The arguments from select_table_range API call.
        :return: List of coordinate dictionaries.
        """
        try:
            # Get Excel range parameters
            start_row = args.get("start_row", 1)
            start_col = args.get("start_col", 1)
            end_row = args.get("end_row", start_row)
            end_col = args.get("end_col", start_col)

            print(f"🎯 Calculating coordinates for Excel range {self._col_num_to_letter(start_col)}{start_row}:{self._col_num_to_letter(end_col)}{end_row}")

            # Method 1 (Priority): Precise detection based on _annotation_dict
            range_cells = self._get_cells_in_range_from_annotation(start_row, start_col, end_row, end_col)
            
            if range_cells:
                # If cells in range are found, calculate their merged rectangle
                merged_rect = self._calculate_merged_rectangle(range_cells)
                if merged_rect:
                    print(f"✅ Using annotation-based coordinates: {merged_rect}")
                    
                    # Convert to relative coordinates of application window
                    app_rect = self.application_window.rectangle()
                    relative_rect = {
                        "left": merged_rect["left"] - app_rect.left,
                        "top": merged_rect["top"] - app_rect.top,
                        "right": merged_rect["right"] - app_rect.left,
                        "bottom": merged_rect["bottom"] - app_rect.top
                    }
                    
                    # Ensure coordinates are within window bounds
                    relative_rect["left"] = max(0, min(relative_rect["left"], app_rect.width() - 1))
                    relative_rect["top"] = max(0, min(relative_rect["top"], app_rect.height() - 1))
                    relative_rect["right"] = max(relative_rect["left"] + 1, min(relative_rect["right"], app_rect.width()))
                    relative_rect["bottom"] = max(relative_rect["top"] + 1, min(relative_rect["bottom"], app_rect.height()))
                    
                    return [relative_rect]

            # Method 2 (Fallback): Traditional method based on landmark detection
            print("⚠️ Falling back to landmark-based coordinate calculation")
            layout = self._detect_excel_interface_layout()
            
            # Calculate the range coordinates relative to the Excel worksheet area
            left = layout['worksheet_left'] + (start_col - 1) * layout['cell_width']
            top = layout['worksheet_top'] + (start_row - 1) * layout['cell_height']
            right = layout['worksheet_left'] + end_col * layout['cell_width']
            bottom = layout['worksheet_top'] + end_row * layout['cell_height']
            
            # Get application window rectangle for bounds checking
            app_rect = self.application_window.rectangle()
            
            # Ensure coordinates are within window bounds
            left = max(0, min(left, app_rect.width() - 1))
            top = max(0, min(top, app_rect.height() - 1))
            right = max(left + 1, min(right, app_rect.width()))
            bottom = max(top + 1, min(bottom, app_rect.height()))

            fallback_rect = {
                "left": float(left),
                "top": float(top),
                "right": float(right),
                "bottom": float(bottom)
            }

            print(f"📍 Using landmark-based coordinates: {fallback_rect}")


            # Return in the format expected by capture_from_adjusted_coords
            return [fallback_rect]
            
        except Exception as e:
            print(f"❌ Error calculating Excel range coordinates: {e}")
            return []

    def handle_screenshot_status(self) -> None:
        """
        Handle the screenshot status when the annotation is overlapped and the agent is unable to select the control items.
        """

        utils.print_with_color(
            "Annotation is overlapped and the agent is unable to select the control items. New annotated screenshot is taken.",
            "magenta",
        )
        self.control_reannotate = self.app_agent.Puppeteer.execute_command(
            "annotation", self._args, self._annotation_dict
        )

    def sync_memory(self):
        """
        Sync the memory of the AppAgent.
        """

        app_root = self.control_inspector.get_application_root_name(
            self.application_window
        )

        action_type = [
            self.app_agent.Puppeteer.get_command_types(action.function)
            for action in self.actions.actions
        ]

        all_previous_success_actions = self.get_all_success_actions()

        action_success = self.actions.to_list_of_dicts(
            success_only=True, previous_actions=all_previous_success_actions
        )

        # Create the additional memory data for the log.
        additional_memory = AppAgentAdditionalMemory(
            Step=self.session_step,
            RoundStep=self.round_step,
            AgentStep=self.app_agent.step,
            Round=self.round_num,
            Subtask=self.subtask,
            SubtaskIndex=self.round_subtask_amount,
            FunctionCall=self.actions.get_function_calls(),
            Action=self.actions.to_list_of_dicts(
                previous_actions=all_previous_success_actions
            ),
            ActionSuccess=action_success,
            ActionType=action_type,
            Request=self.request,
            Agent="AppAgent",
            AgentName=self.app_agent.name,
            Application=app_root,
            Cost=self._cost,
            Results=self.actions.get_results(),
            error=self._exeception_traceback,
            time_cost=self._time_cost,
            ControlLog=self.actions.get_control_logs(),
            UserConfirm=(
                "Yes"
                if self.status.upper()
                == self._agent_status_manager.CONFIRM.value.upper()
                else None
            ),
        )

        # Log the original response from the LLM.
        self.add_to_memory(self._response_json)

        # Log the additional memory data for the AppAgent.
        self.add_to_memory(asdict(additional_memory))

    def update_memory(self) -> None:
        """
        Update the memory of the Agent.
        """

        # Sync the memory of the AppAgent.
        self.sync_memory()

        self.app_agent.add_memory(self._memory_data)

        # Log the memory item.
        self.context.add_to_structural_logs(self._memory_data.to_dict())
        # self.log(self._memory_data.to_dict())

        # Only memorize the keys in the HISTORY_KEYS list to feed into the prompt message in the future steps.
        memorized_action = {
            key: self._memory_data.to_dict().get(key)
            for key in configs.get("HISTORY_KEYS", [])
        }

        if self.is_confirm():

            if self._is_resumed:
                self._memory_data.add_values_from_dict({"UserConfirm": "Yes"})
                memorized_action["UserConfirm"] = "Yes"
            else:
                self._memory_data.add_values_from_dict({"UserConfirm": "No"})
                memorized_action["UserConfirm"] = "No"

        # Save the screenshot to the blackboard if the SaveScreenshot flag is set to True by the AppAgent.
        self._update_image_blackboard()
        self.host_agent.blackboard.add_trajectories(memorized_action)

    def get_all_success_actions(self) -> List[Dict[str, Any]]:
        """
        Get the previous action.
        :return: The previous action of the agent.
        """
        agent_memory = self.app_agent.memory

        if agent_memory.length > 0:
            success_action_memory = agent_memory.filter_memory_from_keys(
                ["ActionSuccess"]
            )
            success_actions = []
            for success_action in success_action_memory:
                success_actions += success_action.get("ActionSuccess", [])

        else:
            success_actions = []

        return success_actions

    def get_last_success_actions(self) -> List[Dict[str, Any]]:
        """
        Get the previous action.
        :return: The previous action of the agent.
        """
        agent_memory = self.app_agent.memory

        if agent_memory.length > 0:
            last_success_actions = (
                agent_memory.get_latest_item().to_dict().get("ActionSuccess", [])
            )

        else:
            last_success_actions = []

        return last_success_actions

    def _update_image_blackboard(self) -> None:
        """
        Save the screenshot to the blackboard if the SaveScreenshot flag is set to True by the AppAgent.
        """
        screenshot_saving = self._response_json.get("SaveScreenshot", {})

        if screenshot_saving.get("save", False):

            screenshot_save_path = self.log_path + f"action_step{self.session_step}.png"
            metadata = {
                "screenshot application": self.context.get(
                    ContextNames.APPLICATION_PROCESS_NAME
                ),
                "saving reason": screenshot_saving.get("reason", ""),
            }
            self.app_agent.blackboard.add_image(screenshot_save_path, metadata)

    def _save_to_xml(self) -> None:
        """
        Save the XML file for the current state. Only work for COM objects.
        """
        log_abs_path = os.path.abspath(self.log_path)
        xml_save_path = os.path.join(
            log_abs_path, f"xml/action_step{self.session_step}.xml"
        )
        self.app_agent.Puppeteer.save_to_xml(xml_save_path)

    def get_filtered_annotation_dict(
        self, annotation_dict: Dict[str, UIAWrapper], configs: Dict[str, Any] = configs
    ) -> Dict[str, UIAWrapper]:
        """
        Get the filtered annotation dictionary.
        :param annotation_dict: The annotation dictionary.
        :return: The filtered annotation dictionary.
        """

        # Get the control filter type and top k plan from the configuration.
        control_filter_type = configs["CONTROL_FILTER_TYPE"]
        topk_plan = configs["CONTROL_FILTER_TOP_K_PLAN"]

        if len(control_filter_type) == 0 or self.prev_plan == []:
            return annotation_dict

        control_filter_type_lower = [
            control_filter_type_lower.lower()
            for control_filter_type_lower in control_filter_type
        ]

        filtered_annotation_dict = {}

        # Get the top k recent plans from the memory.
        plans = self.control_filter_factory.get_plans(self.prev_plan, topk_plan)

        # Filter the annotation dictionary based on the keywords of control text and plan.
        if "text" in control_filter_type_lower:
            model_text = self.control_filter_factory.create_control_filter("text")
            filtered_text_dict = model_text.control_filter(annotation_dict, plans)
            filtered_annotation_dict = (
                self.control_filter_factory.inplace_append_filtered_annotation_dict(
                    filtered_annotation_dict, filtered_text_dict
                )
            )

        # Filter the annotation dictionary based on the semantic similarity of the control text and plan with their embeddings.
        if "semantic" in control_filter_type_lower:
            model_semantic = self.control_filter_factory.create_control_filter(
                "semantic", configs["CONTROL_FILTER_MODEL_SEMANTIC_NAME"]
            )
            filtered_semantic_dict = model_semantic.control_filter(
                annotation_dict, plans, configs["CONTROL_FILTER_TOP_K_SEMANTIC"]
            )
            filtered_annotation_dict = (
                self.control_filter_factory.inplace_append_filtered_annotation_dict(
                    filtered_annotation_dict, filtered_semantic_dict
                )
            )

        # Filter the annotation dictionary based on the icon image icon and plan with their embeddings.
        if "icon" in control_filter_type_lower:
            model_icon = self.control_filter_factory.create_control_filter(
                "icon", configs["CONTROL_FILTER_MODEL_ICON_NAME"]
            )

            cropped_icons_dict = self.photographer.get_cropped_icons_dict(
                self.application_window, annotation_dict
            )
            filtered_icon_dict = model_icon.control_filter(
                annotation_dict,
                cropped_icons_dict,
                plans,
                configs["CONTROL_FILTER_TOP_K_ICON"],
            )
            filtered_annotation_dict = (
                self.control_filter_factory.inplace_append_filtered_annotation_dict(
                    filtered_annotation_dict, filtered_icon_dict
                )
            )

        return filtered_annotation_dict

    def _get_cells_in_range_from_annotation(self, start_row: int, start_col: int, end_row: int, end_col: int) -> List[Dict[str, float]]:
        """
        Find all cell controls within the specified range from _annotation_dict and return their rectangle information.
        This method is more accurate than landmark detection because it directly uses visual recognition results.
        
        :param start_row: Starting row number (1-based)
        :param start_col: Starting column number (1-based)
        :param end_row: Ending row number (1-based)
        :param end_col: Ending column number (1-based)
        :return: List of rectangles for all cells within the range
        """
        if not self._api_annotation_dict:
            print("⚠️ No annotation dictionary available for cell range detection")
            return []
        
        range_cells = []
        found_cells_info = []
        
        print(f"🔍 Looking for cells in range {self._col_num_to_letter(start_col)}{start_row}:{self._col_num_to_letter(end_col)}{end_row}")
        
        # Iterate through all recognized controls
        for control_label, control_element in self._api_annotation_dict.items():
            try:
                # Get control text content and position information
                control_text = getattr(control_element, 'window_text', lambda: '')()
                control_type = getattr(control_element.element_info, 'control_type', '')
                
                # Check if this is a cell type control
                if self._is_cell_control(control_element, control_text):
                    # Try to infer cell position from control location
                    cell_position = self._infer_cell_position(control_element, control_text)
                    cell_row, cell_col = cell_position if cell_position else (None, None)
                    
                    # Check if within target range
                    if (cell_row is not None and cell_col is not None and
                        start_row <= cell_row <= end_row and 
                        start_col <= cell_col <= end_col):
                        
                        rect = control_element.rectangle()
                        cell_info = {
                            "left": float(rect.left),
                            "top": float(rect.top), 
                            "right": float(rect.right),
                            "bottom": float(rect.bottom),
                            "row": float(cell_row),
                            "col": float(cell_col),
                            "label": control_label,
                            "text": control_text
                        }
                        range_cells.append(cell_info)
                        found_cells_info.append(f"{self._col_num_to_letter(cell_col)}{cell_row}({control_label})")
                        
            except Exception as e:
                continue  # Skip controls that cannot be processed
        
        if found_cells_info:
            print(f"✅ Found {len(range_cells)} cells in range: {', '.join(found_cells_info[:10])}" + 
                  (f" and {len(found_cells_info)-10} more..." if len(found_cells_info) > 10 else ""))
        else:
            print("❌ No cells found in the specified range from annotation dictionary")
            
        return range_cells
    
    def _is_cell_control(self, control_element, control_text: str) -> bool:
        """
        Determine if a control is a cell control
        """
        try:
            control_type = getattr(control_element.element_info, 'control_type', '')
            control_class = getattr(control_element.element_info, 'class_name', '')
            source = getattr(control_element.element_info, 'source', '')
            
            # Common cell control characteristics
            cell_indicators = [
                'cell' in control_type.lower(),
                'cell' in control_class.lower(), 
                source == 'grounding',  # Virtual controls from visual recognition
                control_type in ['Edit', 'Text', 'DataItem'],  # Common cell control types
            ]
            
            return any(cell_indicators)
            
        except Exception:
            return False
    
    def _infer_cell_position(self, control_element, control_text: str) -> Optional[tuple]:
        """
        Infer the row and column position of a control in Excel
        :return: (row, col) or None
        """
        try:
            # First try to get cell address from window_text
            actual_text = getattr(control_element, 'window_text', lambda: '')()
            
            # Method 1: If control text is in cell address format (e.g., "A1", "B5")
            text_to_parse = actual_text or control_text
            if text_to_parse and len(text_to_parse) <= 10:  # Reasonable cell address length
                import re
                cell_pattern = r'^([A-Z]+)(\d+)$'
                match = re.match(cell_pattern, text_to_parse.strip().upper())
                if match:
                    col_letter, row_num = match.groups()
                    col_num = self._col_letter_to_num(col_letter)
                    # Note: Use parsed row and column numbers directly, no adjustment needed
                    return (int(row_num), col_num)
            
            # Method 2: Infer based on control's relative position on screen (requires landmark info)
            return self._infer_position_from_coordinates(control_element)
            
        except Exception:
            return None
    
    def _infer_position_from_coordinates(self, control_element) -> Optional[tuple]:
        """
        Infer Excel row and column position based on control's screen coordinates
        :return: (row, col) or None
        """
        try:
            # Use previous landmark detection as fallback
            layout = self._detect_excel_interface_layout()
            rect = control_element.rectangle()
            
            # Calculate offset relative to worksheet start position
            app_rect = self.application_window.rectangle()
            relative_left = rect.left - app_rect.left - layout['worksheet_left']
            relative_top = rect.top - app_rect.top - layout['worksheet_top']
            
            # Calculate row and column numbers (1-based)
            if relative_left >= 0 and relative_top >= 0:
                col = int(relative_left // layout['cell_width']) + 1
                row = int(relative_top // layout['cell_height']) + 1
                return (row, col)
                
        except Exception:
            pass
            
        return None
    
    def _col_letter_to_num(self, col_letter: str) -> int:
        """Convert column letter to number (A=1, B=2, ..., Z=26, AA=27, ...)"""
        num = 0
        for char in col_letter:
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num
    
    def _col_num_to_letter(self, col_num: int) -> str:
        """Convert column number to letter (1=A, 2=B, ..., 26=Z, 27=AA, ...)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    def _calculate_merged_rectangle(self, cells_info: List[Dict[str, float]]) -> Dict[str, float]:
        """
        Calculate merged rectangle area for multiple cells
        """
        if not cells_info:
            return {}
        
        # Find boundaries of all cells
        min_left = min(cell["left"] for cell in cells_info)
        min_top = min(cell["top"] for cell in cells_info) 
        max_right = max(cell["right"] for cell in cells_info)
        max_bottom = max(cell["bottom"] for cell in cells_info)
        
        merged_rect = {
            "left": min_left,
            "top": min_top,
            "right": max_right, 
            "bottom": max_bottom
        }
        
        print(f"📐 Merged rectangle calculated: {merged_rect}")
        return merged_rect
