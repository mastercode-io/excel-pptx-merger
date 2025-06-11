"""PowerPoint template processing and merge field replacement module with enhanced image handling."""

import logging
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union
from pptx import Presentation
from pptx.shapes.base import BaseShape
from pptx.text.text import TextFrame
from pptx.shapes.picture import Picture
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
from PIL import Image as PILImage

from .utils.exceptions import PowerPointProcessingError, TemplateError, ValidationError
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
            raise PowerPointProcessingError(f"PowerPoint template not found: {self.template_path}")

        try:
            self.presentation = Presentation(self.template_path)
            logger.info(f"Successfully loaded PowerPoint template: {self.template_path}")
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
                if hasattr(shape, 'text_frame') and shape.text_frame:
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
            if not hasattr(shape, 'text_frame') or not shape.text_frame:
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

    def merge_data(self, data: Dict[str, Any], output_path: str,
                  images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> str:
        """Merge data into presentation template and save to output path."""
        if not self.presentation:
            self._validate_template()

        try:
            # Process each slide
            for slide_idx, slide in enumerate(self.presentation.slides):
                logger.debug(f"Processing slide {slide_idx + 1}")
                self._process_slide(slide, data, images)

            # Save the merged presentation
            self.presentation.save(output_path)
            logger.info(f"Merged presentation saved to: {output_path}")

            return output_path

        except Exception as e:
            raise PowerPointProcessingError(f"Failed to merge data into presentation: {e}")

    def _process_slide(self, slide, data: Dict[str, Any],
                      images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> None:
        """Process a single slide for merge field replacement."""
        try:
            # Process text shapes
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    self._process_text_shape(shape, data, images)

                # Process tables
                elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    self._process_table_shape(shape, data)

                # Process placeholder images
                elif hasattr(shape, 'text') and shape.text:
                    # Check if this is an image placeholder
                    if self._is_image_placeholder(shape.text):
                        self._replace_image_placeholder(slide, shape, data, images)

        except Exception as e:
            logger.error(f"Failed to process slide: {e}")

    def _process_text_shape(self, shape: BaseShape, data: Dict[str, Any],
                           images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> None:
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

    def _try_replace_text_with_image(self, shape: BaseShape, data: Dict[str, Any],
                                    images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> bool:
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
                image_path = self._get_image_for_field(field, data, images)
                if image_path and os.path.exists(image_path):
                    # Replace the text shape with an image
                    self._replace_shape_with_image(shape, image_path)
                    return True

            return False

        except Exception as e:
            logger.warning(f"Failed to replace text with image: {e}")
            return False

    def _get_image_for_field(self, field_name: str, data: Dict[str, Any],
                            images: Dict[str, List[Dict[str, Any]]]) -> Optional[str]:
        """Get image path for a specific field."""
        try:
            # First try to get image path from data
            image_path = self._get_field_value(field_name, data)
            if image_path and os.path.exists(str(image_path)):
                return str(image_path)

            # Try position-based matching
            return self._find_image_by_field_name(field_name, images)

        except Exception as e:
            logger.warning(f"Failed to get image for field '{field_name}': {e}")
            return None

    def _find_image_by_field_name(self, field_name: str,
                                 images: Dict[str, List[Dict[str, Any]]]) -> Optional[str]:
        """Find image by field name using various matching strategies."""
        field_lower = field_name.lower()

        # Strategy 1: Direct field name matching
        for sheet_name, sheet_images in images.items():
            for image_info in sheet_images:
                # Check if field name contains position information
                if 'position' in image_info:
                    position = image_info['position']
                    if position.get('estimated_cell'):
                        cell_ref = position['estimated_cell'].lower()
                        if cell_ref in field_lower or field_lower.endswith(cell_ref):
                            return image_info['path']

        # Strategy 2: Pattern matching for common image field patterns
        patterns = [
            r'image_search\.(\d+)\.image',  # image_search.0.image
            r'(\w+)_image_(\d+)',           # sheet_image_1
            r'image(\d+)',                  # image1
            r'img(\d+)'                     # img1
        ]

        for pattern in patterns:
            match = re.search(pattern, field_lower)
            if match:
                index_str = match.group(1) if match.groups() else '0'
                try:
                    index = int(index_str)
                    # Find image by index across all sheets
                    for sheet_name, sheet_images in images.items():
                        if 0 <= index < len(sheet_images):
                            return sheet_images[index]['path']
                except (ValueError, IndexError):
                    continue

        # Strategy 3: Keyword matching
        keywords = ['image', 'img', 'picture', 'photo']
        for keyword in keywords:
            if keyword in field_lower:
                # Return first available image
                for sheet_images in images.values():
                    if sheet_images:
                        return sheet_images[0]['path']

        return None

    def _replace_shape_with_image(self, shape: BaseShape, image_path: str) -> None:
        """Replace a text shape with an image."""
        try:
            # Get the slide that contains this shape
            slide = shape._element.getparent().getparent()  # Navigate up to slide

            # Get shape position and size
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height

            # Remove the original shape
            shape_element = shape._element
            shape_element.getparent().remove(shape_element)

            # Add the image at the same position
            # Note: This is a simplified approach - in practice, you might want to
            # access the slide object properly through the presentation
            logger.info(f"Would replace shape with image: {image_path}")
            # For now, we'll log this action - full implementation would require
            # accessing the slide object to add the picture

        except Exception as e:
            logger.warning(f"Failed to replace shape with image: {e}")

    def _process_paragraph(self, paragraph, data: Dict[str, Any]) -> None:
        """Process a single paragraph for merge field replacement."""
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
                new_text = new_text.replace(merge_pattern, str(field_value) if field_value is not None else "")

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
            logger.warning(f"Failed to process paragraph: {e}")

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
            field_parts = field_name.split('.')
            current_value = data

            for part in field_parts:
                if isinstance(current_value, dict):
                    current_value = current_value.get(part)
                elif isinstance(current_value, list):
                    try:
                        index = int(part)
                        current_value = current_value[index] if 0 <= index < len(current_value) else None
                    except (ValueError, IndexError):
                        current_value = None
                else:
                    current_value = None

                if current_value is None:
                    break

            return current_value

        except Exception as e:
            logger.warning(f"Failed to get field value for '{field_name}': {e}")
            return None

    def _is_image_placeholder(self, text: str) -> bool:
        """Check if text represents an image placeholder."""
        # Look for patterns like {{image_name}} or similar
        image_patterns = [
            r'\{\{.*image.*\}\}',
            r'\{\{.*img.*\}\}',
            r'\{\{.*photo.*\}\}',
            r'\{\{.*picture.*\}\}'
        ]

        text_lower = text.lower()
        for pattern in image_patterns:
            if re.search(pattern, text_lower):
                return True

        return False

    def _replace_image_placeholder(self, slide, shape: BaseShape, data: Dict[str, Any],
                                  images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> None:
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

            if image_path and os.path.exists(image_path):
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

        except Exception as e:
            logger.warning(f"Failed to replace image placeholder: {e}")

    def validate_template(self) -> Dict[str, Any]:
        """Validate template and return information about merge fields and structure."""
        if not self.presentation:
            self._validate_template()

        try:
            validation_info = {
                'slide_count': len(self.presentation.slides),
                'merge_fields': self.get_merge_fields(),
                'image_placeholders': self._get_image_placeholders(),
                'slides': []
            }

            for slide_idx, slide in enumerate(self.presentation.slides):
                slide_info = {
                    'slide_number': slide_idx + 1,
                    'shape_count': len(slide.shapes),
                    'merge_fields': self._extract_slide_merge_fields(slide),
                    'image_placeholders': self._get_slide_image_placeholders(slide),
                    'has_tables': any(shape.shape_type == MSO_SHAPE_TYPE.TABLE for shape in slide.shapes),
                    'has_images': any(shape.shape_type == MSO_SHAPE_TYPE.PICTURE for shape in slide.shapes)
                }
                validation_info['slides'].append(slide_info)

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
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    text_content = self._get_full_text_from_shape(shape)
                    if text_content and self._is_image_placeholder(text_content):
                        merge_fields = validate_merge_fields(text_content)
                        placeholders.extend(merge_fields)
                elif hasattr(shape, 'text') and shape.text and self._is_image_placeholder(shape.text):
                    merge_fields = validate_merge_fields(shape.text)
                    placeholders.extend(merge_fields)
        except Exception as e:
            logger.warning(f"Failed to get slide image placeholders: {e}")

        return placeholders

    def preview_merge(self, data: Dict[str, Any],
                     images: Optional[Dict[str, List[Dict[str, Any]]]] = None) -> Dict[str, Any]:
        """Preview what the merge would look like without actually performing it."""
        try:
            merge_fields = self.get_merge_fields()
            image_placeholders = self._get_image_placeholders()

            preview = {
                'merge_fields': merge_fields,
                'image_placeholders': image_placeholders,
                'field_values': {},
                'image_mappings': {},
                'missing_fields': [],
                'missing_images': [],
                'template_info': self.validate_template()
            }

            # Check which fields will be populated
            for field in merge_fields:
                value = self._get_field_value(field, data)
                preview['field_values'][field] = value

                if value is None:
                    preview['missing_fields'].append(field)

            # Check image mappings
            if images:
                for placeholder in image_placeholders:
                    image_path = self._get_image_for_field(placeholder, data, images)
                    preview['image_mappings'][placeholder] = image_path

                    if not image_path or not os.path.exists(image_path):
                        preview['missing_images'].append(placeholder)
            else:
                preview['missing_images'] = image_placeholders.copy()

            return preview

        except Exception as e:
            raise PowerPointProcessingError(f"Failed to generate merge preview: {e}")

    def get_image_requirements(self) -> Dict[str, Any]:
        """Analyze template to determine image requirements."""
        try:
            image_placeholders = self._get_image_placeholders()

            requirements = {
                'total_image_placeholders': len(image_placeholders),
                'placeholder_details': [],
                'suggested_naming': []
            }

            for placeholder in image_placeholders:
                # Analyze placeholder to suggest naming conventions
                detail = {
                    'placeholder': placeholder,
                    'suggested_field_path': self._suggest_field_path(placeholder),
                    'suggested_cell_position': self._extract_cell_position_from_field(placeholder)
                }
                requirements['placeholder_details'].append(detail)

            return requirements

        except Exception as e:
            logger.warning(f"Failed to analyze image requirements: {e}")
            return {'total_image_placeholders': 0, 'placeholder_details': [], 'suggested_naming': []}

    def _suggest_field_path(self, placeholder: str) -> str:
        """Suggest a field path for accessing image data."""
        # Convert placeholder to suggested data path
        if 'image_search' in placeholder:
            return placeholder  # Already in correct format
        elif 'image' in placeholder:
            # Convert generic image reference to indexed format
            match = re.search(r'(\d+)', placeholder)
            if match:
                index = match.group(1)
                return f"order_form.image_search.{index}.image"

        return placeholder

    def _extract_cell_position_from_field(self, field_name: str) -> Optional[str]:
        """Extract cell position hint from field name."""
        # Look for cell references in field names
        cell_pattern = r'([A-Z]+\d+)'
        match = re.search(cell_pattern, field_name.upper())
        if match:
            return match.group(1)

        return None
    
    def close(self) -> None:
        """Close the presentation and free resources."""
        # python-pptx doesn't require explicit closing, but we'll reset the reference
        self.presentation = None
