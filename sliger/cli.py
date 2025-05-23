import ast
import sys
from pathlib import Path
from typing import Any, Optional
from pprint import pprint
import re # Added for imagify

import typer
from google.oauth2 import service_account
from googleapiclient.errors import HttpError # Keep for direct error handling in CLI if any

from . import jinja_utils
from . import slides_utils
from . import drive_utils

app = typer.Typer()

# Global state for the CLI application
# This will hold credentials, presentation ID, Jinja environment, and verbose flag
state: dict[str, Any] = {}


@app.callback()
def main(
    creds_file: Path = typer.Option(
        ..., 
        exists=True, 
        file_okay=True, 
        dir_okay=False, 
        readable=True, 
        resolve_path=True,
        help="Path to the GCP service account credentials JSON file."
    ),
    presentation_id: str = typer.Option(
        ..., 
        help="ID of the Google Slides presentation."
    ),
    config_path: Optional[Path] = typer.Option(
        None, 
        exists=True, 
        file_okay=True, 
        dir_okay=False, 
        readable=True, 
        resolve_path=True,
        help="Optional path to TOML config file for Jinja custom functions."
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Enable verbose output."),
):
    """Sliger: Automate Google Slides with Python and Jinja2."""
    try:
        state["creds"] = service_account.Credentials.from_service_account_file(creds_file)
    except Exception as e:
        typer.echo(f"Error loading credentials from {creds_file}: {e}", err=True)
        raise typer.Exit(code=1)
        
    state["presentation_id"] = presentation_id
    state["verbose"] = verbose
    # Convert config_path to string if it's a Path object, or None
    config_path_str = str(config_path) if config_path else None
    state["jinja_env"] = jinja_utils.load_jinja_environment(config_path_str)

    if verbose:
        typer.echo("State initialized.")
        typer.echo(f"  Credentials loaded: {"Yes" if state.get("creds") else "No"}")
        typer.echo(f"  Presentation ID: {state.get("presentation_id")}")
        typer.echo(f"  Jinja env loaded: {"Yes" if state.get("jinja_env") else "No"}")
        if config_path_str:
            typer.echo(f"  Jinja config path: {config_path_str}")

@app.command()
def duplicate_presentation(copy_title: str = typer.Option(..., help="Title for the new duplicated presentation.")):
    """Duplicates the presentation specified by --presentation-id."""
    presentation_id = state["presentation_id"]
    creds = state["creds"]
    verbose = state["verbose"]

    drive_service = drive_utils.get_drive_service(creds)
    if not drive_service:
        typer.echo("Failed to initialize Google Drive service.", err=True)
        raise typer.Exit(code=1)

    if verbose:
        typer.echo(f"Attempting to duplicate presentation ID: {presentation_id} to '{copy_title}'")

    presentation_copy_id = drive_utils.copy_file(drive_service, presentation_id, copy_title)

    if presentation_copy_id:
        if verbose:
            typer.echo(f"Successfully copied presentation. New ID: {presentation_copy_id}")
        
        permission_id = drive_utils.set_file_permissions_anyone_writer(
            drive_service, presentation_copy_id, verbose
        )
        if not permission_id:
            typer.echo("Failed to set permissions on the new presentation.", err=True)
            # Continue, but inform the user

        final_message = (
            f"Created new presentation: https://docs.google.com/presentation/d/{presentation_copy_id}/edit"
        )
        if verbose:
            typer.echo(final_message)
        else:
            typer.echo(presentation_copy_id) # Output only the ID if not verbose
        # Potentially return this or store in state if other commands need it, though not currently designed for that.
    else:
        typer.echo(f"Failed to duplicate presentation '{presentation_id}'.", err=True)
        raise typer.Exit(code=1)

@app.command()
def delete_slide(slide_number: int = typer.Option(..., help="1-based index of the slide to delete.")):
    """Deletes a specific slide by its number (1-based index)."""
    presentation_id = state["presentation_id"]
    creds = state["creds"]
    verbose = state["verbose"]

    service = slides_utils.get_slides_service(creds)
    if not service:
        typer.echo("Failed to initialize Google Slides service.", err=True)
        raise typer.Exit(code=1)

    slides = slides_utils.get_presentation_slides(service, presentation_id, verbose)
    if not slides:
        # Error already printed by get_presentation_slides or no slides found
        raise typer.Exit(code=1)

    slides_utils.print_slide_details_if_verbose(slides, verbose) # This will print details before attempting delete

    slide_to_delete_id = slides_utils.get_slide_id_by_index(slides, slide_number)

    if slide_to_delete_id is None:
        typer.echo(
            f"Slide number {slide_number} doesn't exist. "
            f"Please input an existing slide number (1 to {len(slides)}).",
            err=True,
        )
        raise typer.Exit(code=1)
    
    if verbose:
        typer.echo(f"Attempting to delete slide #{slide_number} (ID: {slide_to_delete_id})")

    try:
        response = slides_utils.delete_slide_by_id(service, presentation_id, slide_to_delete_id)
        if verbose:
            typer.echo("Slide deletion API call successful.")
            pprint(response) # Using pprint for complex API responses
        typer.echo(f"Slide #{slide_number} deleted successfully.")
    except HttpError as err:
        typer.echo(f"Error deleting slide: {err}", err=True)
        raise typer.Exit(code=1)

@app.command()
def duplicate_slide(slide_number: int = typer.Option(..., help="1-based index of the slide to duplicate.")):
    """Duplicates a specific slide by its number (1-based index)."""
    presentation_id = state["presentation_id"]
    creds = state["creds"]
    verbose = state["verbose"]

    service = slides_utils.get_slides_service(creds)
    if not service:
        typer.echo("Failed to initialize Google Slides service.", err=True)
        raise typer.Exit(code=1)

    slides_before = slides_utils.get_presentation_slides(service, presentation_id, verbose)
    if not slides_before:
        raise typer.Exit(code=1)

    slides_utils.print_slide_details_if_verbose(slides_before, verbose)

    slide_to_duplicate_id = slides_utils.get_slide_id_by_index(slides_before, slide_number)

    if slide_to_duplicate_id is None:
        typer.echo(
            f"Slide number {slide_number} doesn't exist. "
            f"Please input an existing slide number (1 to {len(slides_before)}).",
            err=True,
        )
        raise typer.Exit(code=1)

    if verbose:
        typer.echo(f"Attempting to duplicate slide #{slide_number} (ID: {slide_to_duplicate_id})")

    try:
        response = slides_utils.duplicate_slide_by_id(service, presentation_id, slide_to_duplicate_id)
        if verbose:
            typer.echo("Slide duplication API call successful.")
            pprint(response)
        
        typer.echo(f"Slide #{slide_number} duplicated successfully.")

        if verbose: # Fetch and show updated slide list if verbose
            slides_after = slides_utils.get_presentation_slides(service, presentation_id, False) # verbose=False here to avoid double count message
            if slides_after:
                typer.echo(f"\nPresentation now contains {len(slides_after)} slides.")
                # slides_utils.print_slide_details_if_verbose(slides_after, True) # Optionally print all details again

    except HttpError as err:
        typer.echo(f"Error duplicating slide: {err}", err=True)
        raise typer.Exit(code=1)

@app.command()
def jinjify(data: str = typer.Option("{}", callback=ast.literal_eval, help="JSON string of data for Jinja rendering (e.g., '{\"name\": \"World\"}').")):
    """Processes text in slides through Jinja2 for templating."""
    presentation_id = state["presentation_id"]
    creds = state["creds"]
    verbose = state["verbose"]
    jinja_env = state["jinja_env"]

    service = slides_utils.get_slides_service(creds)
    if not service:
        typer.echo("Failed to initialize Google Slides service.", err=True)
        raise typer.Exit(code=1)

    slides = slides_utils.get_presentation_slides(service, presentation_id, verbose)
    if not slides:
        raise typer.Exit(code=1)

    slides_utils.print_slide_details_if_verbose(slides, verbose)

    typer.echo("Processing slides for Jinja templating...")
    total_changes_made = 0

    for i, slide in enumerate(slides):
        slide_id = slide.get("objectId")
        if verbose:
            typer.echo(f"Processing Slide #{i + 1} ({slide_id})")

        text_elements = slides_utils.get_text_elements_from_slide(slide)
        if not text_elements:
            if verbose:
                typer.echo(f"  No text elements found on Slide #{i+1}.")
            continue # Skip to next slide if no text elements
        
        # Ensure parsed_texts and update_requests are within the loop for each slide
        parsed_texts = [
            slides_utils.gslides_element_to_text(el, slide_id) for el in text_elements
        ]
        update_requests = []

        for text_info in parsed_texts:
            original_text = text_info["text"]
            rendered_text = jinja_utils.render_jinja_in_string(jinja_env, original_text, data)

            if rendered_text != original_text:
                if verbose:
                    typer.echo(f"  Change on slide {i+1} (element {text_info['object_id']}):")
                    typer.echo(f"    Original: '{original_text}'")
                    typer.echo(f"    Rendered: '{rendered_text}'")
                
                update_request = slides_utils.text_update_to_gslides_request({
                    "original_text": original_text,
                    "rendered_text": rendered_text,
                    "page_object_id": text_info["page_object_id"],
                    "object_id": text_info["object_id"]
                })
                update_requests.append(update_request)

        if update_requests: # Only make API call if there are changes for this slide
            try:
                if verbose:
                    typer.echo(f"  Applying {len(update_requests)} text updates to slide #{i+1}...")
                body = {"requests": update_requests}
                service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
                total_changes_made += len(update_requests)
                if verbose:
                    typer.echo(f"  Successfully updated text on slide #{i+1}.")
            except HttpError as err:
                typer.echo(f"Error updating text on slide #{i+1}: {err}", err=True)
        elif verbose:
            typer.echo(f"  No text changes to apply on Slide #{i+1}.")
    # This line must be aligned with the `for` loop, i.e., outside it, but inside the function.
    typer.echo(f"Jinjify processing complete. Total text updates made: {total_changes_made}")

@app.command()
def imagify():
    """Replaces image placeholders like ![image](path) with actual images."""
    presentation_id = state["presentation_id"]
    creds = state["creds"]
    verbose = state["verbose"]
    jinja_env = state["jinja_env"]

    slides_service = slides_utils.get_slides_service(creds)
    if not slides_service:
        typer.echo("Failed to initialize Google Slides service.", err=True)
        raise typer.Exit(code=1)
    
    drive_service = drive_utils.get_drive_service(creds)
    if not drive_service:
        typer.echo("Failed to initialize Google Drive service.", err=True)
        raise typer.Exit(code=1)

    slides = slides_utils.get_presentation_slides(slides_service, presentation_id, verbose)
    if not slides:
        raise typer.Exit(code=1)

    slides_utils.print_slide_details_if_verbose(slides, verbose)
    typer.echo("Processing slides for image replacement...")
    total_images_replaced = 0

    for i, slide in enumerate(slides):
        slide_id = slide.get("objectId") # pageObjectId
        if verbose:
            typer.echo(f"Processing Slide #{i + 1} ({slide_id})")

        text_elements = slides_utils.get_text_elements_from_slide(slide)
        if not text_elements and verbose:
            typer.echo(f"  No text elements found on Slide #{i+1}.")
            continue

        for text_element_shape in text_elements:
            # text_element_shape is a PageElement with shapeType TEXT_BOX
            shape_id = text_element_shape.get("objectId")
            parsed_text_info = slides_utils.gslides_element_to_text(text_element_shape, slide_id)
            original_text = parsed_text_info["text"]

            # Render Jinja in the text, in case the image path itself is templated
            # Pass current slide index (0-based) and slide_id as potential Jinja context
            jinja_data_for_image_path = {"slide_index": i, "slide_id": slide_id}
            rendered_text_content = jinja_utils.render_jinja_in_string(jinja_env, original_text, jinja_data_for_image_path)

            image_match = re.fullmatch(r"!\[image\]\((.*)\)", rendered_text_content.strip())

            if image_match:
                image_path_template = image_match.group(1)
                # The path itself might be a Jinja expression, render it again if it looks like one.
                # This is a simple check; more robust would be to always try rendering.
                if "{{" in image_path_template and "}}" in image_path_template:
                     image_path = jinja_utils.render_jinja_in_string(jinja_env, image_path_template, jinja_data_for_image_path).strip()
                else:
                    image_path = image_path_template.strip()
                
                if not Path(image_path).exists():
                    typer.echo(f"  Image path '{image_path}' (from '{rendered_text_content}') does not exist. Skipping.", err=True)
                    continue

                if verbose:
                    typer.echo(f"  Found image placeholder: '{rendered_text_content}' on slide #{i+1}. Path: '{image_path}'")
                    typer.echo(f"  Uploading image '{image_path}' to Google Drive...")

                drive_image_id = drive_utils.upload_image_to_drive(drive_service, image_path, verbose)
                if not drive_image_id:
                    typer.echo(f"  Failed to upload image '{image_path}' to Drive. Skipping replacement.", err=True)
                    continue
                
                image_url_on_drive = f"https://drive.google.com/uc?id={drive_image_id}"

                # Use the properties of the existing text_element_shape for the new image
                size = text_element_shape.get("size")
                transform = text_element_shape.get("transform")

                requests = [
                    {
                        "createImage": {
                            "url": image_url_on_drive,
                            "elementProperties": {
                                "pageObjectId": slide_id, # slide_id is the pageObjectId
                                "size": size,
                                "transform": transform,
                            },
                        }
                    },
                    # Delete the original text box shape after image is created
                    {"deleteObject": {"objectId": shape_id}}
                ]

                try:
                    if verbose:
                        typer.echo(f"  Replacing text box '{shape_id}' with image from '{image_path}' on slide #{i+1}.")
                    
                    body = {"requests": requests}
                    response = slides_service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
                    
                    create_image_response = response.get("replies")[0].get("createImage")
                    if verbose and create_image_response:
                        typer.echo(f"  Created image with ID: {create_image_response.get('objectId')}")
                    total_images_replaced += 1
                except HttpError as err:
                    typer.echo(f"  Error replacing text box with image on slide #{i+1}: {err}", err=True)
                except (IndexError, TypeError) as e:
                    typer.echo(f"  Error parsing API response after image creation: {e}", err=True)
    
    typer.echo(f"Imagify processing complete. Total images replaced: {total_images_replaced}")


if __name__ == "__main__":
    app() # This allows running cli.py directly for testing, though typical entry is via __main__.py or installed script 