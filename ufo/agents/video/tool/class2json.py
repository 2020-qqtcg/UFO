from pydantic import BaseModel, Field
from enum import Enum
from typing import Optional, Union, get_origin, get_args
import json
import inspect
from typing import get_type_hints, Any, List, Dict, Union


def add_additional_properties_false(schema) -> None:
    """
    Recursively add "additionalProperties": false to all object definitions in a JSON schema.
    """
    if "type" in schema and schema["type"] == "object":
        schema.setdefault("additionalProperties", False)

    # If the schema has properties, recursively process them
    if "properties" in schema:
        for prop in schema["properties"].values():
            add_additional_properties_false(prop)

    # Process definitions in $defs (if present)
    if "$defs" in schema:
        for definition in schema["$defs"].values():
            add_additional_properties_false(definition)

    # Process items if this is an array type with item schemas
    if "items" in schema and isinstance(schema["items"], dict):
        add_additional_properties_false(schema["items"])


def remove_titles_recursively(schema) -> None:
    """
    Recursively remove the "title" property from all objects in a JSON schema.
    """
    # Remove "title" if present
    schema.pop("title", None)

    # If the schema has properties, recursively process them
    if "properties" in schema:
        for prop in schema["properties"].values():
            remove_titles_recursively(prop)

    # Process definitions in $defs (if present)
    if "$defs" in schema:
        for definition in schema["$defs"].values():
            remove_titles_recursively(definition)

    # Process items if this is an array type with item schemas
    if "items" in schema and isinstance(schema["items"], dict):
        remove_titles_recursively(schema["items"])


def generate_json_schema(model) -> dict:
    """
    Generate a JSON schema for a Pydantic model.
    """
    # Ensure forward references are resolved
    model.model_rebuild()
    # Generate the schema
    schema = model.model_json_schema()

    add_additional_properties_false(schema)
    schema['additionalProperties'] = False
    remove_titles_recursively(schema)

    # Return the schema
    return schema


class ExcelOperation(BaseModel):
    output: str
    change_location: str
    reason: str

from pydantic import BaseModel, Field
from typing import Dict


class Step(BaseModel):  # 必须继承 BaseModel
    title: str
    voiceover_script: str

class Step2(BaseModel):  # 必须继承 BaseModel
    title: str
    written_explanation: str


class VideoStructure(BaseModel):
    judge: bool
    video_title: str
    thematic_opening_line: str
    sample_opening_line: str
    steps: Dict[str, Step] = Field(default_factory=dict)

class VideoStructure0709(BaseModel):
    video_title: str
    thematic_opening_line: str
    steps: Dict[str, Step] = Field(default_factory=dict)

class JudgeStructure(BaseModel):
    judge: bool

class JudgeStructurenew(BaseModel):
    judge: bool
    reason: str

class JudgeStructure2(BaseModel):
    Clear: bool
    Conciseness: bool
    Completeness: bool
    Sequential_Order: bool


class JudgeStructure3(BaseModel):
    Clear: bool
    Conciseness: bool
    Completeness: bool
    Sequential_Order: bool

class DocumentStructure(BaseModel):
    task_title: str
    steps: Dict[str, Step2] = Field(default_factory=dict)


class TitleStructure(BaseModel):
    video_title: str
    thematic_opening_line: str
    judge: bool
    subvideo_title: str
    sample_opening_line: str

if __name__ == "__main__":
    jsonfile = generate_json_schema(JudgeStructure2)
    print(jsonfile)
    file_path = "../data/steps_schema_questionnaire.json"
    with open(file_path, "w") as file:
        json.dump(jsonfile, file, indent=4)

    # JudgeStructurenew
    jsonfile = generate_json_schema(JudgeStructurenew)
    print(jsonfile)
    file_path = "../data/steps_schema_judge_new.json"
    with open(file_path, "w") as file:
        json.dump(jsonfile, file, indent=4)
