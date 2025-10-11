from pydantic import BaseModel, Field
import json
from typing import Optional


# --- (Helper functions, no modification needed) ---

def add_additional_properties_false(schema) -> None:
    """
    Recursively adds "additionalProperties": false to all object definitions in a JSON schema.
    """
    if "type" in schema and schema["type"] == "object":
        schema.setdefault("additionalProperties", False)
    if "properties" in schema:
        for prop in schema["properties"].values():
            add_additional_properties_false(prop)
    if "$defs" in schema:
        for definition in schema["$defs"].values():
            add_additional_properties_false(definition)
    if "items" in schema and isinstance(schema["items"], dict):
        add_additional_properties_false(schema["items"])


def remove_titles_recursively(schema) -> None:
    """
    Recursively removes the "title" property from all objects in a JSON schema.
    """
    schema.pop("title", None)
    if "properties" in schema:
        for prop in schema["properties"].values():
            remove_titles_recursively(prop)
    if "$defs" in schema:
        for definition in schema["$defs"].values():
            remove_titles_recursively(definition)
    if "items" in schema and isinstance(schema["items"], dict):
        remove_titles_recursively(schema["items"])


def generate_json_schema(model) -> dict:
    """
    Generates a JSON schema for a Pydantic model.
    """
    model.model_rebuild(force=True)
    schema = model.model_json_schema()
    add_additional_properties_false(schema)
    schema['additionalProperties'] = False
    remove_titles_recursively(schema)
    return schema


# --- Code Modification Section ---

# 1. Add a shared model to represent "score and reason".
class ScoreWithReason(BaseModel):
    """Represents a rating item that includes a score and a reason."""
    score: int = Field(
        ...,
        ge=1,
        le=5,
        description="A score from 1 to 5."
    )
    reason: str = Field(
        ...,
        description="The specific reason for the given score."
    )


# 2. Modify the main model so that each field uses the new ScoreWithReason model.
class EvaluationScores(BaseModel):
    """
    Defines a standardized JSON output format where each rating dimension
    includes a score and a reason.
    """
    clarity: ScoreWithReason
    conciseness: ScoreWithReason
    correctness: ScoreWithReason
    completeness: ScoreWithReason
    sequential_order: ScoreWithReason
    # Use 'alias' to handle JSON keys with hyphens.
    text_image_mapping: ScoreWithReason


if __name__ == "__main__":
    jsonfile = generate_json_schema(EvaluationScores)

    print("--- Generated JSON Schema ---")
    print(json.dumps(jsonfile, indent=2))

    file_path = "../data/steps_schema_questionnaire_score_object.json"
    with open(file_path, "w", encoding='utf-8') as file:
        json.dump(jsonfile, file, indent=4)

    print(f"\n✅ Schema has been successfully generated and saved to {file_path}")