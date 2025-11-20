# src/formatting_engine.py

from typing import Dict, Any


def apply_formatting(doc_state: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    """
    Applies a set of formatting rules to a document state dictionary.

    Args:
        doc_state (Dict[str, Any]): The DocumentModel JSON to be modified.
        rules (Dict[str, Any]): The formatting rules from the user.

    Returns:
        Dict[str, Any]: The modified DocumentModel JSON.
    """

    # Example Rule Structure for `rules`:
    # {
    #   "style_map": {
    #     "Heading 1": { "font_name": "Arial", "font_size": 16, "bold": true },
    #     "Normal": { "font_name": "Calibri", "font_size": 11 }
    #   },
    #   "global_paragraph_properties": {
    #     "alignment": "left",
    #     "line_spacing": 1.5
    #   }
    # }

    style_map = rules.get("style_map", {})
    global_paragraph_properties = rules.get("global_paragraph_properties", {})

    for section in doc_state.get("sections", []):
        for element in section.get("elements", []):
            if element.get("type") == "paragraph":
                # Ensure properties dictionary exists
                if "properties" not in element or element["properties"] is None:
                    element["properties"] = {}

                # 1. Apply global properties first as a baseline
                element["properties"].update(global_paragraph_properties)

                # 2. Apply style-specific properties, which can override global ones
                style_name = element.get("properties", {}).get("style", "Normal")
                if style_name in style_map:
                    element["properties"].update(style_map[style_name])

    # TODO: Add logic for other element types (e.g., tables) and other rule types.

    return doc_state