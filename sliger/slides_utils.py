import sys
from typing import Optional, Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def get_slides_service(creds: Any):
    """Builds and returns the Google Slides API service."""
    return build("slides", "v1", credentials=creds)


def get_presentation_slides(service: Any, presentation_id: str, verbose: bool) -> Optional[list]:
    """Fetches and returns the slides of a presentation."""
    try:
        presentation = service.presentations().get(presentationId=presentation_id).execute()
        slides = presentation.get("slides")
        if verbose:
            print(f"The presentation contains {len(slides)} slides:")
        return slides
    except HttpError as err:
        print(err, file=sys.stderr)
        return None


def get_slide_id_by_index(slides: list, index: int) -> Optional[str]:
    """Returns the objectId of a slide given its 1-based index."""
    if 0 < index <= len(slides):
        return slides[index - 1].get("objectId")
    return None


def print_slide_details_if_verbose(slides: list, verbose: bool):
    """Prints details of each slide if verbose mode is enabled."""
    if verbose:
        for i, slide in enumerate(slides):
            slide_id = slide.get("objectId")
            page_elements = slide.get("pageElements", [])
            print(
                f"- Slide #{i + 1} ({slide_id}) contains {len(page_elements)} elements."
            )


def get_text_elements_from_slide(slide: dict) -> list:
    """Extracts text box elements from a slide."""
    elements = slide.get("pageElements", [])
    return list(filter(
        lambda el: el.get("shape")
        and el["shape"].get("shapeType") == "TEXT_BOX",
        elements,
    ))


def gslides_element_to_text(el: dict, page_object_id: str) -> dict:
    """Converts a Google Slides text element to a simpler text dictionary.

    Args:
        el: The page element dictionary from Google Slides API.
        page_object_id: The ID of the slide (page) the element is on.

    Returns:
        A dictionary with 'object_id' (of the text element itself),
        'page_object_id', and 'text'.
    """
    shape_object_id = el.get("objectId") # ID of the shape itself
    if "text" not in el["shape"]:
        return {"object_id": shape_object_id, "page_object_id": page_object_id, "text": ""}

    text_contents = map(
        lambda text_el: text_el["textRun"]["content"] if "textRun" in text_el else "",
        el["shape"]["text"]["textElements"],
    )
    text = "".join(text_contents)
    return {"object_id": shape_object_id, "page_object_id": page_object_id, "text": text.strip()}


def text_update_to_gslides_request(change: dict) -> dict:
    """Converts a text change dictionary to a Google Slides API request.

    Args:
        change: A dictionary containing 'original_text', 'rendered_text',
                'page_object_id' (slide ID), and 'object_id' (shape ID).

    Returns:
        A Google Slides API request dictionary for replaceAllText.
    """
    return {
        "replaceAllText": {
            "containsText": {"text": change["original_text"], "matchCase": True},
            "replaceText": change["rendered_text"],
            "pageObjectIds": [change["page_object_id"]], # Targets the whole slide
            # To target a specific shape, you'd use 'replaceAllShapesWithText' and provide the shape's objectId.
            # This current structure assumes replacement across any text box on the specified slide
            # that matches 'containsText'. If specific shape targeting is needed,
            # the request structure and likely the calling logic needs to change.
        }
    }


# Low-level operations
def delete_slide_by_id(service: Any, presentation_id: str, slide_id: str):
    """Deletes a slide by its ID."""
    requests = [{"deleteObject": {"objectId": slide_id}}]
    body = {"requests": requests}
    response = (
        service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    return response


def duplicate_slide_by_id(service: Any, presentation_id: str, slide_id: str):
    """Duplicates a slide by its ID."""
    requests = [{"duplicateObject": {"objectId": slide_id}}]
    body = {"requests": requests}
    response = (
        service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    return response

def delete_element_by_id(service: Any, presentation_id: str, object_id: str):
    """Deletes a page element (e.g., a shape) by its ID."""
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": [{"deleteObject": {"objectId": object_id}}]},
    ).execute()
