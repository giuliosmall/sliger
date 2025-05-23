"""Sliger: Slide of the tiger - Pythonize Google Slides."""

# This file can be used to define the public API of the sliger package.
# For now, as it's primarily a CLI tool, we keep it minimal.
# The CLI application is defined in sliger.cli

__version__ = "0.1.0" # Example version

import ast
import importlib
import re
import sys
import time
from pathlib import Path
from pprint import pprint
from typing import Callable, Optional

import toml
import typer
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from jinja2 import BaseLoader, Environment

app = typer.Typer()
state = {"verbose": False}


# StringLoader is a custom Jinja2 template loader for rendering templates
# provided as strings, instead of loading from files or other sources.
class StringLoader(BaseLoader):
    # get_source() returns a tuple (template, None, lambda: True) for Jinja2
    # to load the template source, without a filename/path, and always up to date.
    def get_source(self, environment, template):
        return template, None, lambda: True


def load_jinja_environment(config_path: str) -> Environment:
    if config_path:
        with open(config_path) as f:
            config = toml.load(f)

        def import_function(function_name: str) -> Callable:
            if "." in function_name:
                mod_name, func_name = function_name.rsplit(".", 1)
                mod = importlib.import_module(mod_name)
                res = getattr(mod, func_name)
            else:
                res = getattr(globals(), function_name)
            return res

        function_map = config["function_map"]
        function_map = {
            name: import_function(function_name)
            for name, function_name in function_map.items()
        }
    else:
        function_map = {}

    env = Environment(loader=StringLoader())
    env.globals.update(function_map)

    return env


def _get_slides_service():
    """Builds and returns the Google Slides API service."""
    return build("slides", "v1", credentials=state["creds"])


def _get_presentation_slides(service, presentation_id: str) -> Optional[list]:
    """Fetches and returns the slides of a presentation."""
    try:
        presentation = service.presentations().get(presentationId=presentation_id).execute()
        slides = presentation.get("slides")
        # The verbose printing will be handled by the calling functions
        # if state["verbose"]:
        #     print(f"The presentation contains {len(slides)} slides:")
        return slides
    except HttpError as err:
        print(err, file=sys.stderr)
        return None


@app.command()
def duplicate_presentation(copy_title: str = typer.Option(...)):
    presentation_id = state["presentation_id"]
    try:
        drive_service = build("drive", "v3", credentials=state["creds"])
        body = {"name": copy_title}
        drive_response = (
            drive_service.files().copy(fileId=presentation_id, body=body).execute()
        )
        presentation_copy_id = drive_response.get("id")

        # Simpler permission setting
        user_permission = {
            "type": "anyone",
            "role": "writer",
        }
        permission_response = drive_service.permissions().create(
            fileId=presentation_copy_id,
            body=user_permission,
            fields="id",
        ).execute()

        if state["verbose"]:
            print(f"Permission Id: {permission_response.get('id')}")

        if state["verbose"]:
            print(
                f"Created new presentation at https://docs.google.com/presentation/d/{presentation_copy_id}/edit"
            )
        else:
            print(presentation_copy_id)

        return presentation_copy_id

    except HttpError as err:
        print(err, file=sys.stderr)

    return None


def delete_slide_by_id(service, presentation_id: str, slide_id: str):
    requests = [{"deleteObject": {"objectId": slide_id}}]

    body = {"requests": requests}
    response = (
        service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    return response


def _get_slide_id_by_index(slides: list, index: int) -> Optional[str]:
    """Returns the objectId of a slide given its 1-based index."""
    if 0 < index <= len(slides):
        return slides[index - 1].get("objectId")
    return None


def _print_slide_details_if_verbose(slides: list):
    """Prints details of each slide if verbose mode is enabled."""
    if state["verbose"]:
        print(f"The presentation contains {len(slides)} slides:")
        for i, slide in enumerate(slides):
            slide_id = slide.get("objectId")
            page_elements = slide.get("pageElements", []) # Ensure pageElements exists
            print(
                f"- Slide #{i + 1} ({slide_id}) contains {len(page_elements)} elements."
            )


@app.command()
def delete_slide(id: int = typer.Option(...)):
    slide_to_delete = id
    presentation_id = state["presentation_id"]
    try:
        service = _get_slides_service()
        slides = _get_presentation_slides(service, presentation_id)

        if not slides:
            return

        _print_slide_details_if_verbose(slides)

        slide_to_delete_id = _get_slide_id_by_index(slides, slide_to_delete)

        if slide_to_delete_id is None:
            print(
                f"\nSlide number {id} doesn't exist. "
                f"Please input an existing slide from your presentation.",
                file=sys.stderr,
            )
        else:
            res = delete_slide_by_id(service, presentation_id, slide_to_delete_id)

            if state["verbose"]:
                pprint(res)

    except HttpError as err:
        print(err)


def duplicate_slide_by_id(service, presentation_id: str, slide_id: str):
    requests = [{"duplicateObject": {"objectId": slide_id}}]

    body = {"requests": requests}
    response = (
        service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    return response


@app.command()
def duplicate_slide(id: int = typer.Option(...)):
    slide_to_duplicate = id
    presentation_id = state["presentation_id"]
    try:
        service = _get_slides_service()
        slides = _get_presentation_slides(service, presentation_id)

        if not slides:
            return

        _print_slide_details_if_verbose(slides)

        slide_to_duplicate_id = _get_slide_id_by_index(slides, slide_to_duplicate)

        if slide_to_duplicate_id is None:
            if state["verbose"]:
                print(
                    f"\nSlide number {id} doesn't exist. "
                    f"Please input an existing slide from your presentation.",
                    file=sys.stderr,
                )
        else:
            res = duplicate_slide_by_id(service, presentation_id, slide_to_duplicate_id)
            if state["verbose"]:
                pprint(res)

            # Call the Slides API again
            presentation = (
                service.presentations().get(presentationId=presentation_id).execute()
            )
            slides = presentation.get("slides")

            if state["verbose"]:
                print("\n Now the presentation contains {} slides".format(len(slides)))

    except HttpError as err:
        print(err)


def gslides_element_to_text(el: dict, object_id: str) -> dict:
    # An object that is of text type but doesn't have any text in it
    if "text" not in el["shape"]:
        return {"object_id": object_id, "text": ""}

    text_contents = map(
        lambda el: el["textRun"]["content"] if "textRun" in el else "",
        el["shape"]["text"]["textElements"],
    )

    text = "".join(text_contents)

    return {"object_id": object_id, "text": text.strip()}


def render_jinja_in_string(string: str, data: Optional[dict] = None) -> str:
    if data is None:
        data = {}

    env = state["jinja_env"]
    rtemplate = env.from_string(string)

    return rtemplate.render(**data)


def strftime_with_ordinal(string: str, t) -> str:
    def ordinal(n: int) -> str:
        """
        derive the ordinal numeral for a given number n
        """
        return f"{n:d}{'tsnrhtdd'[(n // 10 % 10 != 1) * (n % 10 < 4) * n % 10::4]}"

    string = string.replace("%O", ordinal(t.tm_mday))
    return time.strftime(string, t)


def text_update_to_gslides_request(change: dict) -> dict:
    return {
        "replaceAllText": {
            "containsText": {"text": change["text"], "matchCase": True},
            "replaceText": change["rendered_text"],
            "pageObjectIds": [change["object_id"]],
        }
    }


def _get_text_elements_from_slide(slide: dict) -> list:
    """Extracts text box elements from a slide."""
    elements = slide.get("pageElements", [])
    return list(filter(
        lambda el: el.get("shape")
        and el["shape"].get("shapeType") == "TEXT_BOX",
        elements,
    ))


@app.command()
def jinjify(data: str = typer.Option("{}", callback=ast.literal_eval)):
    presentation_id = state["presentation_id"]
    try:
        service = _get_slides_service()
        slides = _get_presentation_slides(service, presentation_id)

        if not slides:
            return

        _print_slide_details_if_verbose(slides)

        for i, slide in enumerate(slides):
            slide_id = slide.get("objectId")

            text_elements = _get_text_elements_from_slide(slide)

            texts = [gslides_element_to_text(el, slide_id) for el in text_elements]
            updated_texts = []
            for text in texts:
                rendered_text = render_jinja_in_string(
                    text["text"],
                    {
                        "now": time.localtime(),
                        "strftime": strftime_with_ordinal,
                        **data,
                    },
                )
                if rendered_text != text["text"]:
                    text["rendered_text"] = rendered_text
                    updated_texts.append(text)

            gslides_update_requests = [
                text_update_to_gslides_request(update) for update in updated_texts
            ]

            if len(gslides_update_requests) != 0:
                if state["verbose"]:
                    print("Executing changes", gslides_update_requests)

                body = {"requests": gslides_update_requests}

                response = (
                    service.presentations()
                    .batchUpdate(presentationId=state["presentation_id"], body=body)
                    .execute()
                )
                print(response)

    except HttpError as err:
        print(err, file=sys.stderr)


def upload_image(
    service,
    drive_service,
    presentation_id: str,
    slide_id: str,
    size=None,
    transform=None,
    img_path=None,
):
    if size is None:
        size = {
            "height": {"magnitude": 405, "unit": "PT"},
            "width": {"magnitude": 720, "unit": "PT"},
        }
    if transform is None:
        transform = {
            "scaleX": 1,
            "scaleY": 1,
            "translateX": 0,
            "translateY": 0,
            "unit": "PT",
        }

    try:
        # Upload the image to Google Drive
        media = MediaFileUpload(img_path, mimetype="image/jpeg", resumable=True)

        img_file = (
            drive_service.files()
            .create(
                body={"name": img_path, "parents": ["root"]},
                media_body=media,
                fields="id",
            )
            .execute()
        )

        # Make the image publicly accessible
        permission = {
            "type": "anyone",
            "role": "reader",
        }
        drive_service.permissions().create(
            fileId=img_file.get("id"), body=permission
        ).execute()

        # Get the presentation
        presentation = (
            service.presentations().get(presentationId=presentation_id).execute()
        )
        slides = presentation.get("slides")

        # Select the target slide
        page_id = slides[slide_id]["objectId"]

        # Define the request to add the image at the center of the slide.
        requests = [
            {
                "createImage": {
                    "url": "https://drive.google.com/uc?id=" + img_file.get("id"),
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": size,
                        "transform": transform,
                    },
                }
            }
        ]

        # Execute the request.
        response = (
            service.presentations()
            .batchUpdate(presentationId=presentation_id, body={"requests": requests})
            .execute()
        )

        create_image_response = response.get("replies")[0].get("createImage")

        if state["verbose"]:
            print(
                "Created image with ID: {0}".format(
                    create_image_response.get("objectId")
                )
            )

    except HttpError as error:
        print(f"An error occurred: {error}", file=sys.stderr)


def delete_element(service, presentation_id: str, object_id: str):
    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": [{"deleteObject": {"objectId": object_id}}]},
    ).execute()


@app.command()
def imagify():
    presentation_id = state["presentation_id"]
    try:
        service = _get_slides_service()
        drive_service = build("drive", "v3", credentials=state["creds"])
        slides = _get_presentation_slides(service, presentation_id)

        if not slides:
            return

        _print_slide_details_if_verbose(slides)

        for i, slide in enumerate(slides):
            slide_id = slide.get("objectId")

            text_elements = _get_text_elements_from_slide(slide)

            for text_element in text_elements:
                text = gslides_element_to_text(text_element, slide_id)
                text = render_jinja_in_string(
                    text["text"],
                    {"now": time.localtime(), "strftime": strftime_with_ordinal},
                )
                image_match = re.fullmatch(r"!\[image\]\((.*)\)", text)
                if image_match:
                    image_path = image_match.group(1)
                    upload_image(
                        service,
                        drive_service,
                        presentation_id,
                        i,
                        text_element["size"],
                        text_element["transform"],
                        image_path,
                    )
                    delete_element(service, presentation_id, text_element["objectId"])

    except HttpError as err:
        print(err, file=sys.stderr)


@app.callback()
def main(
    creds_file: Path = typer.Option(...),
    presentation_id: str = typer.Option(...),
    config_path: str = typer.Option(None),
    verbose: bool = False,
):
    state["creds"] = service_account.Credentials.from_service_account_file(creds_file)
    state["presentation_id"] = presentation_id
    state["jinja_env"] = load_jinja_environment(config_path)
    state["verbose"] = verbose


if __name__ == "__main__":
    app()
