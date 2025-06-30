"""PowerPoint template processing and merge field replacement module."""

import logging
import os
import re
from typing import Any, Dict, List, Optional
from PIL import Image as PILImage
from pptx import Presentation
from pptx.shapes.base import BaseShape
from pptx.enum.shapes import MSO_SHAPE_TYPE
from .utils.exceptions import PowerPointProcessingError
from .utils.validation import validate_merge_fields

logger = logging.getLogger(__name__)


class PowerPointProcessor:
    """Processes PowerPoint templates and replaces merge fields with data."""

    def __init__(self, template_path: str) -> None:
        """Initialize PowerPoint processor with template path."""
        self.template_path = template_path
        self.presentation = None
        self._validate_template()

    def _validate_template(self) -> None:
        """Validate PowerPoint template exists and is readable."""
        if not os.path.exists(self.template_path):
            raise PowerPointProcessingError(
                f"PowerPoint template not found: {self.template_path}"
            )

        try:
            self.presentation = Presentation(self.template_path)
            logger.info(
                f"Successfully loaded PowerPoint template: {self.template_path}"
            )
        except Exception as e:
            raise PowerPointProcessingError(f"Invalid PowerPoint template format: {e}")

    def get_merge_fields(self) -> List[str]:
        """Extract all merge fields from the presentation."""
        if not self.presentation:
            self._validate_template()

        merge_fields = set()

        try:
            for slide in self.presentation.slides:
                slide_fields = self._extract_slide_merge_fields(slide)
                merge_fields.update(slide_fields)

            return sorted(list(merge_fields))

        except Exception as e:
            raise PowerPointProcessingError(f"Failed to extract merge fields: {e}")

    def _extract_slide_merge_fields(self, slide) -> List[str]:
        """Extract merge fields from a single slide."""
        fields = []

        try:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    text_content = self._get_full_text_from_shape(shape)
                    if text_content:
                        shape_fields = validate_merge_fields(text_content)
                        fields.extend(shape_fields)

                # Check for table cells
                if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    table_fields = self._extract_table_merge_fields(shape.table)
                    fields.extend(table_fields)

        except Exception as e:
            logger.warning(f"Failed to extract merge fields from slide: {e}")

        return fields

    def _extract_table_merge_fields(self, table) -> List[str]:
        """Extract merge fields from table cells."""
        fields = []

        try:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text:
                        cell_fields = validate_merge_fields(cell.text)
                        fields.extend(cell_fields)
        except Exception as e:
            logger.warning(f"Failed to extract merge fields from table: {e}")

        return fields

    def _get_full_text_from_shape(self, shape: BaseShape) -> str:
        """Get complete text content from a shape including all paragraphs."""
        try:
            if not hasattr(shape, "text_frame") or not shape.text_frame:
                return ""

            text_parts = []
            for paragraph in shape.text_frame.paragraphs:
                paragraph_text = ""
                for run in paragraph.runs:
                    paragraph_text += run.text
                text_parts.append(paragraph_text)

            return "\n".join(text_parts)

        except Exception as e:
            logger.warning(f"Failed to get text from shape: {e}")
            return ""

    def merge_data(
        self,
        data: Dict[str, Any],
        output_path: str,
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> str:
        """Merge data into the PowerPoint template and save to output path.

        Args:
            data: Data to merge into the template
            output_path: Path to save the merged presentation
            images: Dictionary of images by sheet name

        Returns:
            Path to the merged presentation
        """
        if not self.presentation:
            raise PowerPointProcessingError("No presentation loaded")

        # Ensure output_path is an absolute path
        if not os.path.isabs(output_path):
            output_path = os.path.abspath(output_path)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)
        logger.debug(f"Ensuring PowerPoint output directory exists: {output_dir}")

        try:
            # Process each slide
            for slide_idx, slide in enumerate(self.presentation.slides):
                logger.debug(f"Processing slide {slide_idx + 1}")
                self._process_slide(slide, data, images)

            # Validate and clean up before saving
            self._validate_presentation_integrity()
            
            # Save the merged presentation
            self.presentation.save(output_path)
            logger.info(f"Merged presentation saved to: {output_path}")

            return output_path

        except Exception as e:
            raise PowerPointProcessingError(
                f"Failed to merge data into presentation: {e}"
            )

    def _process_slide(
        self,
        slide,
        data: Dict[str, Any],
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> None:
        """Process a single slide for merge field replacement."""
        try:
            # Process text shapes
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    self._process_text_shape(shape, data, images)

                # Process tables
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    self._process_table_shape(shape, data)

                # Process placeholder images
                elif hasattr(shape, "text") and shape.text:
                    # Check if this is an image placeholder
                    if self._is_image_placeholder(shape.text):
                        self._replace_image_placeholder(slide, shape, data, images)

        except Exception as e:
            logger.error(f"Failed to process slide: {e}")

    def _process_text_shape(
        self,
        shape: BaseShape,
        data: Dict[str, Any],
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> None:
        """Process text shape for merge field replacement."""
        try:
            if not shape.text_frame:
                return

            # First check if this might be an image placeholder
            shape_text = self._get_full_text_from_shape(shape)
            if shape_text and self._is_image_placeholder(shape_text):
                # Try to replace with image first
                if self._try_replace_text_with_image(shape, data, images):
                    return  # Successfully replaced with image, skip text processing

            # Process each paragraph for text replacement
            for paragraph in shape.text_frame.paragraphs:
                self._process_paragraph(paragraph, data)

        except Exception as e:
            logger.warning(f"Failed to process text shape: {e}")

    def _try_replace_text_with_image(
        self,
        shape: BaseShape,
        data: Dict[str, Any],
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> bool:
        """Try to replace text shape with image if it's an image placeholder."""
        if not images:
            return False

        try:
            shape_text = self._get_full_text_from_shape(shape)
            merge_fields = validate_merge_fields(shape_text)

            if not merge_fields:
                return False

            # Try to find image for the first merge field
            for field in merge_fields:
                # Check if this is an image field
                if self._is_image_field(field, data):
                    image_path = self._get_image_for_field(field, data, images)
                    if image_path:
                        # Verify the image file exists
                        if os.path.exists(str(image_path)):
                            # Replace the text shape with an image
                            return self._replace_shape_with_image(shape, image_path)
                        else:
                            logger.warning(
                                f"Image file does not exist at path: {image_path}"
                            )
                            # Keep the original text if image insertion fails
                            return False

            return False

        except Exception as e:
            logger.warning(f"Failed to replace text with image: {e}")
            return False

    def _get_field_type(self, field_name: str, data: Dict[str, Any]) -> str:
        """Get the type of a field from metadata.

        Args:
            field_name: The field name or path to check
            data: The data structure containing field values and metadata

        Returns:
            The field type as a string ('text', 'image', etc.)
        """
        # Check for direct field type in the current data level
        if "_field_types" in data and field_name in data["_field_types"]:
            return data["_field_types"][field_name]

        # For nested fields, try to navigate to the parent object
        parts = field_name.split(".")
        if len(parts) > 1:
            # Try to find the parent object
            current_data = data
            parent_path = ".".join(parts[:-1])
            leaf_name = parts[-1]

            # First try direct path
            parent = self._get_field_value(parent_path, data)

            if parent and isinstance(parent, dict):
                if "_field_types" in parent and leaf_name in parent["_field_types"]:
                    return parent["_field_types"][leaf_name]

            # Try to find in sheet data
            for sheet_name, sheet_data in data.items():
                if isinstance(sheet_data, dict):
                    # Try to find the field in this sheet
                    for table_name, table_data in sheet_data.items():
                        if isinstance(table_data, list):
                            # Check each row in the table
                            for row in table_data:
                                if (
                                    isinstance(row, dict)
                                    and "_field_types" in row
                                    and leaf_name in row["_field_types"]
                                ):
                                    return row["_field_types"][leaf_name]

        # Default to text if no type information is found
        return "text"

    def _is_image_field(self, field_name: str, data: Dict[str, Any]) -> bool:
        """Determine if a field is an image field based on metadata."""
        field_type = self._get_field_type(field_name, data)
        return field_type == "image"

    def _is_image_placeholder(self, text_content: str) -> bool:
        """Check if text content is an image placeholder.

        Now uses field type information when available, falling back to
        heuristic detection only when necessary.
        """
        if not text_content:
            return False

        # Extract merge fields
        merge_fields = validate_merge_fields(text_content)
        if not merge_fields:
            return False

        # We'll check the field type when we have the data
        # For now, just return True if it looks like an image placeholder
        # This will be refined when we have the actual data and field types
        return any("image" in field.lower() for field in merge_fields)

    def _get_image_for_field(
        self,
        field_name: str,
        data: Dict[str, Any],
        images: Dict[str, List[Dict[str, Any]]],
    ) -> Optional[str]:
        """Get image path for a specific field."""
        try:
            # Check if this is an image field based on metadata
            is_image = self._is_image_field(field_name, data)

            # First try to get image data from data
            image_data = self._get_field_value(field_name, data)
            if image_data:
                if is_image:
                    logger.debug(
                        f"Found image data for field '{field_name}': {type(image_data)}"
                    )

                    # Handle new mixed data structure (dict with base64/path)
                    if isinstance(image_data, dict):
                        # Try path first if it exists and is valid
                        if "path" in image_data and image_data["path"]:
                            path = str(image_data["path"])
                            if os.path.exists(path):
                                logger.debug(f"Using existing file path: {path}")
                                return path

                        # Fall back to base64 data
                        if "base64" in image_data and image_data["base64"]:
                            # Create temporary file from base64 data
                            temp_path = self._create_temp_image_from_base64(
                                image_data["base64"]
                            )
                            if temp_path:
                                logger.debug(
                                    f"Created temporary image from base64: {temp_path}"
                                )
                                return temp_path

                    # Handle legacy string path format
                    elif isinstance(image_data, str):
                        if os.path.exists(str(image_data)):
                            logger.debug(f"Using image file path: {image_data}")
                            return str(image_data)
                        else:
                            logger.warning(
                                f"Image file does not exist at path: {image_data}"
                            )
                else:
                    logger.debug(
                        f"Field '{field_name}' is not an image field, but has value: {type(image_data)}"
                    )
                    # Not an image field, don't return the path
                    return None

            # If field is an image but we didn't find a path, try position-based matching
            if is_image:
                image_path = self._find_image_by_field_name(field_name, images)
                if image_path:
                    # Verify the image file exists
                    if os.path.exists(str(image_path)):
                        logger.debug(
                            f"Found image by field name matching: {image_path}"
                        )
                        return str(image_path)
                    else:
                        logger.warning(
                            f"Image file from field name matching does not exist: {image_path}"
                        )

            # Log if no image was found for an image field
            if is_image:
                logger.warning(f"No image found for image field: {field_name}")
                # Debug log the available images
                if images:
                    logger.debug(f"Available images: {images}")
                else:
                    logger.warning("No images available")

            return None

        except Exception as e:
            logger.warning(f"Failed to get image for field '{field_name}': {e}")
            return None

    def _find_image_by_field_name(
        self, field_name: str, images: Dict[str, List[Dict[str, Any]]]
    ) -> Optional[str]:
        """Find image by field name using various matching strategies."""
        if not images:
            logger.warning("No images provided to _find_image_by_field_name")
            return None

        logger.debug(f"Looking for image matching field: {field_name}")

        # Log available images for debugging
        for sheet_name, sheet_images in images.items():
            logger.debug(f"Sheet {sheet_name} has {len(sheet_images)} images")
            for idx, img in enumerate(sheet_images):
                logger.debug(
                    f"  Image {idx}: {img.get('filename')} at {img.get('path')}"
                )

        field_lower = field_name.lower()

        # Strategy 1: Direct field name matching
        for sheet_name, sheet_images in images.items():
            for image_info in sheet_images:
                # Check if field name contains position information
                if "position" in image_info:
                    position = image_info["position"]
                    if position.get("estimated_cell"):
                        cell_ref = position["estimated_cell"].lower()
                        if cell_ref in field_lower or field_lower.endswith(cell_ref):
                            logger.info(
                                f"Found image by position match: {image_info['path']}"
                            )
                            return image_info["path"]

        # Strategy 2: Pattern matching for common image field patterns
        patterns = [
            r"image_search\.(\d+)\.image",  # image_search.0.image
            r"(\w+)_image_(\d+)",  # sheet_image_1
            r"image(\d+)",  # image1
            r"img(\d+)",  # img1
        ]

        for pattern in patterns:
            match = re.search(pattern, field_lower)
            if match:
                index_str = match.group(1) if match.groups() else "0"
                try:
                    index = int(index_str)
                    # Find image by index across all sheets
                    for sheet_name, sheet_images in images.items():
                        if 0 <= index < len(sheet_images):
                            logger.info(
                                f"Found image by pattern match: {sheet_images[index]['path']}"
                            )
                            return sheet_images[index]["path"]
                except (ValueError, IndexError):
                    continue

        # Strategy 3: Keyword matching
        keywords = ["image", "img", "picture", "photo"]
        for keyword in keywords:
            if keyword in field_lower:
                # Return first available image
                for sheet_name, sheet_images in images.values():
                    if sheet_images:
                        logger.info(
                            f"Found image by keyword match: {sheet_images[0]['path']}"
                        )
                        return sheet_images[0]["path"]

        # Strategy 4: Just use the first available image if all else fails
        for sheet_name, sheet_images in images.items():
            if sheet_images:
                logger.info(
                    f"No specific match found, using first available image: {sheet_images[0]['path']}"
                )
                return sheet_images[0]["path"]

        logger.warning(f"No image found for field: {field_name}")
        return None

    def _replace_shape_with_image(self, shape, image_path: str) -> bool:
        """Replace a shape with an image while maintaining position and size.

        Args:
            shape: The shape to replace
            image_path: Path to the image file

        Returns:
            True if successful, False otherwise
        """
        try:
            # Verify the image file exists
            if not os.path.exists(image_path):
                logger.error(f"Image file does not exist: {image_path}")
                return False

            # Ensure image_path is absolute
            if not os.path.isabs(image_path):
                image_path = os.path.abspath(image_path)

            # Get the parent slide
            slide = shape.part.slide

            # Get the shape dimensions and position
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height

            # Store shape properties before removal
            shape_name = getattr(shape, 'name', 'Unknown')
            
            # Clear the text from the shape to prepare for replacement
            if hasattr(shape, "text_frame") and shape.text_frame:
                try:
                    shape.text_frame.clear()  # Use proper API method
                    logger.debug(f"Cleared text from shape '{shape_name}' before replacement")
                except Exception as text_err:
                    logger.warning(f"Could not clear text from shape '{shape_name}': {text_err}")

            # Use python-pptx API for safer shape removal
            try:
                # Get slide reference
                slide_shapes = slide.shapes
                
                # Find and remove the shape using the proper API
                for i, slide_shape in enumerate(slide_shapes):
                    if slide_shape == shape:
                        # Remove using the shape collection API (safer than direct XML)
                        del slide_shapes[i]
                        logger.debug(f"Safely removed shape '{shape_name}' from slide")
                        break
                        
            except Exception as removal_err:
                logger.warning(f"Could not remove shape using API, trying XML method: {removal_err}")
                
                # Fallback to XML method if API fails
                try:
                    shape_id = shape.shape_id
                    sp_tree = slide.shapes._spTree
                    
                    for sp in sp_tree.findall(".//{*}sp"):
                        if sp.get("id") == str(shape_id):
                            sp_tree.remove(sp)
                            logger.debug(f"Removed shape '{shape_name}' using XML fallback")
                            break
                            
                except Exception as xml_err:
                    logger.error(f"Failed to remove shape '{shape_name}': {xml_err}")
                    return False

            # Add the image while maintaining aspect ratio
            try:
                # Get image dimensions
                with PILImage.open(image_path) as img:
                    img_width, img_height = img.size

                # Calculate aspect ratios
                shape_ratio = width / height
                img_ratio = img_width / img_height

                # Adjust dimensions to maintain aspect ratio
                if img_ratio > shape_ratio:
                    # Image is wider than shape
                    new_width = width
                    new_height = width / img_ratio
                    new_top = top + (height - new_height) / 2
                    new_left = left
                else:
                    # Image is taller than shape
                    new_height = height
                    new_width = height * img_ratio
                    new_left = left + (width - new_width) / 2
                    new_top = top

                # Add the image to the slide
                picture = slide.shapes.add_picture(
                    image_path, new_left, new_top, new_width, new_height
                )

                # Ensure the image is on top (higher z-order)
                # This is done by moving the picture to the end of the shapes collection
                try:
                    pic_element = None
                    for shape_xml in sp_tree.findall(".//{*}pic"):
                        if shape_xml.get("id") == str(picture.shape_id):
                            pic_element = shape_xml
                            break

                    if pic_element is not None and pic_element.getparent() is not None:
                        # Remove and re-append to put it at the end (top of z-order)
                        parent = pic_element.getparent()
                        parent.remove(pic_element)
                        parent.append(pic_element)
                        logger.debug(f"Adjusted z-order to ensure image is on top")
                except Exception as z_err:
                    logger.warning(f"Could not adjust z-order: {z_err}")

                logger.info(f"Successfully replaced shape with image: {image_path}")
                return True

            except Exception as e:
                logger.error(f"Error adding image to slide: {e}")

                # Add a text box with error message as fallback
                tb = slide.shapes.add_textbox(left, top, width, height)
                tf = tb.text_frame
                tf.text = f"Image Error: {os.path.basename(image_path)}"

                return False

        except Exception as e:
            logger.error(f"Error replacing shape with image: {e}")
            return False

    def _process_paragraph(self, paragraph, data: Dict[str, Any]) -> None:
        """Process a single paragraph for merge field replacement while preserving formatting."""
        try:
            # Try the new formatting-preserving approach first
            if self._process_paragraph_preserve_formatting(paragraph, data):
                return

            # Fall back to the original approach if the new one fails
            logger.warning("Falling back to original paragraph processing")
            self._process_paragraph_original(paragraph, data)

        except Exception as e:
            logger.warning(f"Failed to process paragraph: {e}")

    def _process_paragraph_preserve_formatting(
        self, paragraph, data: Dict[str, Any]
    ) -> bool:
        """Process paragraph while preserving run-level formatting."""
        try:
            # Find all merge fields and their positions within runs
            field_positions = self._find_merge_fields_in_runs(paragraph)

            if not field_positions:
                return True  # No merge fields to process

            # Process each merge field
            for field_info in field_positions:
                field_name = field_info["field"]
                field_value = self._get_field_value(field_name, data)
                field_value_str = str(field_value) if field_value is not None else ""

                # Replace the field in the runs
                self._replace_field_in_runs(paragraph, field_info, field_value_str)
            
            # Clean up any empty runs that might cause PowerPoint issues
            self._cleanup_empty_runs(paragraph)

            return True

        except Exception as e:
            logger.warning(
                f"Failed to process paragraph with formatting preservation: {e}"
            )
            return False

    def _process_paragraph_original(self, paragraph, data: Dict[str, Any]) -> None:
        """Original paragraph processing method (fallback)."""
        try:
            # Get the full paragraph text
            paragraph_text = ""
            for run in paragraph.runs:
                paragraph_text += run.text

            # Find merge fields in paragraph
            merge_fields = validate_merge_fields(paragraph_text)

            if not merge_fields:
                return

            # Replace merge fields
            new_text = paragraph_text
            for field in merge_fields:
                field_value = self._get_field_value(field, data)
                merge_pattern = f"{{{{{field}}}}}"
                new_text = new_text.replace(
                    merge_pattern, str(field_value) if field_value is not None else ""
                )

            # Update paragraph text if it changed
            if new_text != paragraph_text:
                # Clear existing runs
                for run in paragraph.runs:
                    run.text = ""

                # Set new text in first run
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
                else:
                    paragraph.add_run().text = new_text

        except Exception as e:
            logger.warning(f"Failed to process paragraph with original method: {e}")

    def _find_merge_fields_in_runs(self, paragraph) -> List[Dict[str, Any]]:
        """Find merge fields and their positions within paragraph runs."""
        field_positions = []

        try:
            # Build a map of text positions to runs
            run_map = []
            text_position = 0

            for run_idx, run in enumerate(paragraph.runs):
                run_text = run.text
                run_start = text_position
                run_end = text_position + len(run_text)

                run_map.append(
                    {
                        "run_idx": run_idx,
                        "run": run,
                        "text": run_text,
                        "start": run_start,
                        "end": run_end,
                    }
                )

                text_position = run_end

            # Get full paragraph text
            full_text = "".join(run["text"] for run in run_map)

            # Find all merge fields in the full text
            merge_fields = validate_merge_fields(full_text)

            for field in merge_fields:
                field_pattern = f"{{{{{field}}}}}"
                field_start = full_text.find(field_pattern)

                if field_start != -1:
                    field_end = field_start + len(field_pattern)

                    # Find which runs contain this field
                    affected_runs = []
                    for run_info in run_map:
                        # Check if this run overlaps with the field
                        if (
                            run_info["start"] < field_end
                            and run_info["end"] > field_start
                        ):
                            # Calculate the portion of the field in this run
                            field_start_in_run = max(0, field_start - run_info["start"])
                            field_end_in_run = min(
                                len(run_info["text"]), field_end - run_info["start"]
                            )

                            affected_runs.append(
                                {
                                    "run_idx": run_info["run_idx"],
                                    "run": run_info["run"],
                                    "field_start_in_run": field_start_in_run,
                                    "field_end_in_run": field_end_in_run,
                                    "run_text": run_info["text"],
                                }
                            )

                    if affected_runs:
                        field_positions.append(
                            {
                                "field": field,
                                "field_pattern": field_pattern,
                                "field_start": field_start,
                                "field_end": field_end,
                                "affected_runs": affected_runs,
                            }
                        )

            return field_positions

        except Exception as e:
            logger.warning(f"Failed to find merge fields in runs: {e}")
            return []

    def _replace_field_in_runs(
        self, paragraph, field_info: Dict[str, Any], replacement_text: str
    ) -> None:
        """Replace a merge field in runs while preserving formatting."""
        try:
            affected_runs = field_info["affected_runs"]

            if len(affected_runs) == 1:
                # Simple case: field is entirely within one run
                self._replace_field_in_single_run(affected_runs[0], replacement_text)
            else:
                # Complex case: field spans multiple runs
                self._replace_field_across_runs(
                    paragraph, affected_runs, replacement_text
                )

        except Exception as e:
            logger.warning(f"Failed to replace field in runs: {e}")

    def _replace_field_in_single_run(
        self, run_info: Dict[str, Any], replacement_text: str
    ) -> None:
        """Replace field within a single run."""
        try:
            run = run_info["run"]
            original_text = run_info["run_text"]
            field_start = run_info["field_start_in_run"]
            field_end = run_info["field_end_in_run"]

            # Build new text by replacing the field portion
            new_text = (
                original_text[:field_start]
                + replacement_text
                + original_text[field_end:]
            )

            # Update the run text (formatting is preserved automatically)
            run.text = new_text
            logger.debug(
                f"Replaced field in single run: '{original_text}' -> '{new_text}'"
            )

        except Exception as e:
            logger.warning(f"Failed to replace field in single run: {e}")

    def _replace_field_across_runs(
        self, paragraph, affected_runs: List[Dict[str, Any]], replacement_text: str
    ) -> None:
        """Replace field that spans across multiple runs."""
        try:
            if not affected_runs:
                return

            # Strategy: Put the replacement text in the first affected run and clear field portions from others
            first_run_info = affected_runs[0]
            first_run = first_run_info["run"]

            # Build the replacement text for the first run
            original_text = first_run_info["run_text"]
            field_start = first_run_info["field_start_in_run"]

            # Keep text before the field + replacement text
            new_first_run_text = original_text[:field_start] + replacement_text

            # Add any text after the field from the last run
            if len(affected_runs) > 1:
                last_run_info = affected_runs[-1]
                last_run_text = last_run_info["run_text"]
                field_end_in_last = last_run_info["field_end_in_run"]

                # Add remaining text from last run
                new_first_run_text += last_run_text[field_end_in_last:]
            else:
                # Single run case (shouldn't happen here, but handle it)
                field_end = first_run_info["field_end_in_run"]
                new_first_run_text += original_text[field_end:]

            # Update first run with complete replacement
            first_run.text = new_first_run_text

            # Clear field portions from other affected runs
            for i in range(1, len(affected_runs)):
                run_info = affected_runs[i]
                run = run_info["run"]
                original_text = run_info["run_text"]

                if i == len(affected_runs) - 1:
                    # Last run: keep text after the field
                    field_end = run_info["field_end_in_run"]
                    run.text = original_text[field_end:]
                else:
                    # Middle runs: completely consumed by the field
                    # Don't create empty runs - remove them instead
                    if run.text.strip():  # Only clear if there was actual content
                        run.text = ""

            logger.debug(f"Replaced field across {len(affected_runs)} runs")

        except Exception as e:
            logger.warning(f"Failed to replace field across runs: {e}")

    def _preserve_run_formatting(self, source_run, target_run) -> None:
        """Copy formatting properties from source run to target run."""
        try:
            # Copy font properties
            if hasattr(source_run, "font") and hasattr(target_run, "font"):
                source_font = source_run.font
                target_font = target_run.font

                # Copy common font properties
                if source_font.name:
                    target_font.name = source_font.name
                if source_font.size:
                    target_font.size = source_font.size
                if source_font.bold is not None:
                    target_font.bold = source_font.bold
                if source_font.italic is not None:
                    target_font.italic = source_font.italic
                if source_font.underline is not None:
                    target_font.underline = source_font.underline
                if source_font.color.rgb:
                    target_font.color.rgb = source_font.color.rgb

                logger.debug("Copied font formatting between runs")

        except Exception as e:
            logger.warning(f"Failed to preserve run formatting: {e}")

    def _cleanup_empty_runs(self, paragraph) -> None:
        """Remove empty runs that can cause PowerPoint repair issues."""
        try:
            runs_to_remove = []
            
            for run in paragraph.runs:
                # Check if run is truly empty (no text or only whitespace)
                if not run.text or not run.text.strip():
                    # Check if run has meaningful formatting that should be preserved
                    has_formatting = False
                    if hasattr(run, 'font'):
                        font = run.font
                        if (font.bold is not None or font.italic is not None or 
                            font.underline is not None or font.size is not None or
                            font.name is not None):
                            has_formatting = True
                    
                    # Only remove if it's truly empty with no special formatting
                    if not has_formatting:
                        runs_to_remove.append(run)
            
            # Remove empty runs (but keep at least one run in the paragraph)
            if len(runs_to_remove) < len(paragraph.runs):
                for run in runs_to_remove:
                    try:
                        # Remove the run's XML element from the paragraph
                        paragraph._element.remove(run._element)
                        logger.debug("Removed empty run from paragraph")
                    except Exception as remove_err:
                        logger.warning(f"Could not remove empty run: {remove_err}")
            
        except Exception as e:
            logger.warning(f"Failed to cleanup empty runs: {e}")

    def _validate_presentation_integrity(self) -> None:
        """Validate presentation integrity to prevent PowerPoint repair issues."""
        try:
            if not self.presentation:
                return
                
            issues_found = 0
            
            for slide_idx, slide in enumerate(self.presentation.slides):
                try:
                    # Check each shape on the slide
                    for shape_idx, shape in enumerate(slide.shapes):
                        # Check text frames for empty paragraphs/runs
                        if hasattr(shape, 'text_frame') and shape.text_frame:
                            for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                                # Ensure paragraph has at least one run
                                if len(paragraph.runs) == 0:
                                    # Add a minimal run to prevent issues
                                    paragraph.add_run("")
                                    issues_found += 1
                                    logger.debug(f"Added empty run to paragraph in slide {slide_idx + 1}, shape {shape_idx}")
                                
                                # Check for completely empty paragraphs
                                total_text = ''.join(run.text for run in paragraph.runs)
                                if not total_text and len(paragraph.runs) > 1:
                                    # Multiple empty runs - consolidate to one
                                    for run in paragraph.runs[1:]:
                                        try:
                                            paragraph._element.remove(run._element)
                                        except:
                                            pass
                                    issues_found += 1
                                    logger.debug(f"Consolidated empty runs in slide {slide_idx + 1}, shape {shape_idx}")
                
                except Exception as shape_err:
                    logger.warning(f"Error validating slide {slide_idx + 1}: {shape_err}")
            
            if issues_found > 0:
                logger.info(f"Fixed {issues_found} potential PowerPoint compatibility issues")
            else:
                logger.debug("No presentation integrity issues found")
                
        except Exception as e:
            logger.warning(f"Failed to validate presentation integrity: {e}")

    def _process_table_shape(self, shape: BaseShape, data: Dict[str, Any]) -> None:
        """Process table shape for merge field replacement."""
        try:
            table = shape.table

            for row in table.rows:
                for cell in row.cells:
                    if cell.text:
                        # Process each paragraph in the cell
                        for paragraph in cell.text_frame.paragraphs:
                            self._process_paragraph(paragraph, data)

        except Exception as e:
            logger.warning(f"Failed to process table shape: {e}")

    def _get_field_value(self, field_name: str, data: Dict[str, Any]) -> Any:
        """Get value for a merge field from data dictionary."""
        try:
            # Handle nested field references like "table.0.field_name"
            field_parts = field_name.split(".")
            current_value = data

            # Debug logging
            logger.debug(f"Getting field value for: {field_name}")
            logger.debug(f"Data keys at root level: {list(data.keys())}")

            # First try direct path resolution
            for part in field_parts:
                if isinstance(current_value, dict):
                    current_value = current_value.get(part)
                    logger.debug(f"After part '{part}': {type(current_value).__name__}")
                elif isinstance(current_value, list):
                    try:
                        index = int(part)
                        current_value = (
                            current_value[index]
                            if 0 <= index < len(current_value)
                            else None
                        )
                        logger.debug(
                            f"After list index {index}: {type(current_value).__name__}"
                        )
                    except (ValueError, IndexError):
                        current_value = None
                        logger.debug(f"Invalid list index: {part}")
                else:
                    current_value = None
                    logger.debug(
                        f"Cannot navigate further from {type(current_value).__name__}"
                    )

                if current_value is None:
                    break

            # If direct path failed and we have sheet data, try looking in each sheet
            if current_value is None:
                logger.debug(
                    "Direct path resolution failed, trying sheet-nested lookup"
                )
                # Try to find the field in sheet data (e.g., order_form.image_search.0.field)
                for sheet_name, sheet_data in data.items():
                    # Skip metadata and debug fields
                    if sheet_name.startswith("__"):
                        continue

                    # Check if this sheet contains the first part of our field path
                    if isinstance(sheet_data, dict) and field_parts[0] in sheet_data:
                        logger.debug(f"Found {field_parts[0]} in sheet {sheet_name}")
                        # Start with the sheet data
                        nested_value = sheet_data

                        # Navigate through the field parts
                        for part in field_parts:
                            if isinstance(nested_value, dict):
                                nested_value = nested_value.get(part)
                            elif isinstance(nested_value, list):
                                try:
                                    index = int(part)
                                    nested_value = (
                                        nested_value[index]
                                        if 0 <= index < len(nested_value)
                                        else None
                                    )
                                except (ValueError, IndexError):
                                    nested_value = None
                            else:
                                nested_value = None

                            if nested_value is None:
                                break

                        if nested_value is not None:
                            logger.debug(
                                f"Found value via sheet-nested lookup: {nested_value}"
                            )
                            return nested_value

                # Special handling for flat structures without row index
                # For fields like "client_info.client_name" (without row index)
                # Only apply this fallback if the field doesn't contain a numeric index
                if len(field_parts) >= 2:
                    table_name = field_parts[0]
                    field_key = field_parts[-1]
                    
                    # Check if any part is a numeric index - if so, don't use fallback
                    has_numeric_index = False
                    for part in field_parts[1:-1]:  # Check middle parts for numeric indices
                        try:
                            int(part)
                            has_numeric_index = True
                            break
                        except ValueError:
                            continue
                    
                    # Only apply fallback for non-indexed fields
                    if not has_numeric_index:
                        # Look for the table in each sheet
                        for sheet_name, sheet_data in data.items():
                            if isinstance(sheet_data, dict) and table_name in sheet_data:
                                table_data = sheet_data[table_name]

                                # Case 1: Table is a flat dictionary (key-value pairs)
                                if isinstance(table_data, dict) and field_key in table_data:
                                    value = table_data[field_key]
                                    logger.debug(
                                        f"Found value in flat structure {table_name}.{field_key}: {value}"
                                    )
                                    return value

                                # Case 2: Table is a list with a single item (only for non-indexed fields)
                                elif isinstance(table_data, list) and len(table_data) > 0:
                                    # Try first row if no index specified
                                    first_row = table_data[0]
                                    if (
                                        isinstance(first_row, dict)
                                        and field_key in first_row
                                    ):
                                        value = first_row[field_key]
                                        logger.debug(
                                            f"Found value in first row of {table_name}[0].{field_key}: {value}"
                                        )
                                        return value

            logger.debug(f"Final value: {current_value}")
            return current_value

        except Exception as e:
            logger.warning(f"Failed to get field value for '{field_name}': {e}")
            return None

    def _replace_image_placeholder(
        self,
        slide,
        shape: BaseShape,
        data: Dict[str, Any],
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> None:
        """Replace image placeholder with actual image."""
        if not images:
            return

        try:
            # Extract image field name from placeholder text
            placeholder_text = shape.text
            merge_fields = validate_merge_fields(placeholder_text)

            if not merge_fields:
                return

            # Get the first merge field (assuming single image placeholder)
            image_field = merge_fields[0]
            image_path = self._get_image_for_field(image_field, data, images)

            if image_path:
                # Verify the image file exists
                if os.path.exists(str(image_path)):
                    try:
                        # Get shape position and size
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        # Remove the placeholder shape
                        slide.shapes._spTree.remove(shape._element)

                        # Add the image
                        slide.shapes.add_picture(image_path, left, top, width, height)
                        logger.debug(f"Replaced image placeholder with: {image_path}")
                    except Exception as img_error:
                        logger.warning(
                            f"Failed to add image '{image_path}': {img_error}"
                        )
                        # Keep the original placeholder if image insertion fails
                        pass
                else:
                    logger.warning(f"Image file does not exist at path: {image_path}")

        except Exception as e:
            logger.warning(f"Failed to replace image placeholder: {e}")

    def preview_merge(
        self,
        data: Dict[str, Any],
        images: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> Dict[str, Any]:
        """Preview what the merge would look like without actually performing it."""
        try:
            merge_fields = self.get_merge_fields()
            image_placeholders = self._get_image_placeholders()

            preview = {
                "merge_fields": merge_fields,
                "image_placeholders": image_placeholders,
                "field_values": {},
                "image_mappings": {},
                "missing_fields": [],
                "missing_images": [],
            }

            # Check which fields will be populated
            for field in merge_fields:
                value = self._get_field_value(field, data)
                preview["field_values"][field] = value

                if value is None:
                    preview["missing_fields"].append(field)

            # Check image mappings
            if images:
                for placeholder in image_placeholders:
                    image_path = self._get_image_for_field(placeholder, data, images)
                    preview["image_mappings"][placeholder] = image_path

                    if not image_path:
                        preview["missing_images"].append(placeholder)
            else:
                preview["missing_images"] = image_placeholders.copy()

            return preview

        except Exception as e:
            raise PowerPointProcessingError(f"Failed to generate merge preview: {e}")

    def get_image_requirements(self) -> Dict[str, Any]:
        """Analyze template to determine image requirements."""
        try:
            image_placeholders = self._get_image_placeholders()

            requirements = {
                "total_image_placeholders": len(image_placeholders),
                "placeholder_details": [],
                "suggested_naming": [],
            }

            for placeholder in image_placeholders:
                # Analyze placeholder to suggest naming conventions
                detail = {
                    "placeholder": placeholder,
                    "suggested_field_path": self._suggest_field_path(placeholder),
                    "suggested_cell_position": self._extract_cell_position_from_field(
                        placeholder
                    ),
                }
                requirements["placeholder_details"].append(detail)

            return requirements

        except Exception as e:
            logger.warning(f"Failed to analyze image requirements: {e}")
            return {
                "total_image_placeholders": 0,
                "placeholder_details": [],
                "suggested_naming": [],
            }

    def _suggest_field_path(self, placeholder: str) -> str:
        """Suggest a field path for accessing image data."""
        # Convert placeholder to suggested data path
        if "image_search" in placeholder:
            return placeholder  # Already in correct format
        elif "image" in placeholder:
            # Convert generic image reference to indexed format
            match = re.search(r"(\d+)", placeholder)
            if match:
                index = match.group(1)
                return f"order_form.image_search.{index}.image"

        return placeholder

    def _extract_cell_position_from_field(self, field_name: str) -> Optional[str]:
        """Extract cell position hint from field name."""
        # Look for cell references in field names
        cell_pattern = r"([A-Z]+\d+)"
        match = re.search(cell_pattern, field_name.upper())
        if match:
            return match.group(1)

        return None

    def _create_temp_image_from_base64(self, base64_data: str) -> Optional[str]:
        """Create a temporary image file from base64 data."""
        try:
            import base64
            import tempfile
            import os

            # Extract the image format and data
            if base64_data.startswith("data:image/"):
                # Format: data:image/png;base64,iVBORw0KGgo...
                header, encoded_data = base64_data.split(",", 1)
                image_format = header.split("/")[1].split(";")[
                    0
                ]  # Extract 'png' from 'data:image/png;base64'
            else:
                # Raw base64 data, assume PNG
                encoded_data = base64_data
                image_format = "png"

            # Decode base64 data
            image_bytes = base64.b64decode(encoded_data)

            # Create temporary file
            with tempfile.NamedTemporaryFile(
                suffix=f".{image_format}", delete=False
            ) as temp_file:
                temp_file.write(image_bytes)
                temp_path = temp_file.name

            logger.debug(f"Created temporary image file: {temp_path}")
            return temp_path

        except Exception as e:
            logger.error(f"Failed to create temporary image from base64: {e}")
            return None

    def close(self) -> None:
        """Close the presentation and free resources."""
        # python-pptx doesn't require explicit closing, but we'll reset the reference
        self.presentation = None

    def validate_template(self) -> Dict[str, Any]:
        """Validate template and return information about merge fields and structure."""
        if not self.presentation:
            self._validate_template()

        try:
            validation_info = {
                "slide_count": len(self.presentation.slides),
                "merge_fields": self.get_merge_fields(),
                "image_placeholders": self._get_image_placeholders(),
                "slides": [],
            }

            for slide_idx, slide in enumerate(self.presentation.slides):
                slide_info = {
                    "slide_number": slide_idx + 1,
                    "shape_count": len(slide.shapes),
                    "merge_fields": self._extract_slide_merge_fields(slide),
                    "image_placeholders": self._get_slide_image_placeholders(slide),
                    "has_tables": any(
                        shape.shape_type == MSO_SHAPE_TYPE.TABLE
                        for shape in slide.shapes
                    ),
                    "has_images": any(
                        shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                        for shape in slide.shapes
                    ),
                }
                validation_info["slides"].append(slide_info)

            return validation_info

        except Exception as e:
            raise PowerPointProcessingError(f"Template validation failed: {e}")

    def _get_image_placeholders(self) -> List[str]:
        """Get all image placeholders from the presentation."""
        placeholders = []

        try:
            for slide in self.presentation.slides:
                slide_placeholders = self._get_slide_image_placeholders(slide)
                placeholders.extend(slide_placeholders)
        except Exception as e:
            logger.warning(f"Failed to get image placeholders: {e}")

        return list(set(placeholders))  # Remove duplicates

    def _get_slide_image_placeholders(self, slide) -> List[str]:
        """Get image placeholders from a single slide."""
        placeholders = []

        try:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    text_content = self._get_full_text_from_shape(shape)
                    if text_content and self._is_image_placeholder(text_content):
                        merge_fields = validate_merge_fields(text_content)
                        placeholders.extend(merge_fields)
                elif (
                    hasattr(shape, "text")
                    and shape.text
                    and self._is_image_placeholder(shape.text)
                ):
                    merge_fields = validate_merge_fields(shape.text)
                    placeholders.extend(merge_fields)
        except Exception as e:
            logger.warning(f"Failed to get slide image placeholders: {e}")

        return placeholders
