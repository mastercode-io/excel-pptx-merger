#!/usr/bin/env python3
"""
Fixed version of the merge field replacement methods.

This addresses the critical bug in multi-run field replacement where fields
spanning multiple runs were not being replaced correctly.
"""

import logging
import re
from typing import Any, Dict, List

logger = logging.getLogger(__name__)

def validate_merge_fields(template_text: str) -> List[str]:
    """Extract and validate merge fields from template text."""
    merge_field_pattern = r"\{\{([^}]+)\}\}"
    fields = re.findall(merge_field_pattern, template_text)
    
    validated_fields = []
    for field in fields:
        field = field.strip()
        if field:
            validated_fields.append(field)
    
    return validated_fields

class FixedPowerPointProcessor:
    """Fixed version of PowerPoint processor with corrected merge field replacement."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def _process_paragraph_preserve_formatting(
        self, paragraph, data: Dict[str, Any]
    ) -> bool:
        """Process paragraph while preserving run-level formatting - FIXED VERSION."""
        try:
            # FIXED: Process fields in reverse order to avoid position shifts
            # This prevents earlier replacements from affecting later field positions
            field_positions = self._find_merge_fields_in_runs(paragraph)

            if not field_positions:
                return True  # No merge fields to process

            # FIXED: Sort fields by their start position in reverse order
            # This ensures we process from right to left, preventing position shifts
            field_positions.sort(key=lambda x: x["field_start"], reverse=True)

            # Process each merge field from right to left
            for field_info in field_positions:
                field_name = field_info["field"]
                field_value = self._get_field_value(field_name, data)
                field_value_str = str(field_value) if field_value is not None else ""

                # Replace the field in the runs
                self._replace_field_in_runs_fixed(paragraph, field_info, field_value_str)

            # Clean up any empty runs that might cause PowerPoint issues
            self._cleanup_empty_runs(paragraph)

            return True

        except Exception as e:
            logger.warning(
                f"Failed to process paragraph with formatting preservation: {e}"
            )
            return False

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

    def _replace_field_in_runs_fixed(
        self, paragraph, field_info: Dict[str, Any], replacement_text: str
    ) -> None:
        """FIXED: Replace a merge field in runs while preserving formatting."""
        try:
            affected_runs = field_info["affected_runs"]

            if len(affected_runs) == 1:
                # Simple case: field is entirely within one run
                self._replace_field_in_single_run(affected_runs[0], replacement_text)
            else:
                # Complex case: field spans multiple runs - FIXED VERSION
                self._replace_field_across_runs_fixed(
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

        except Exception as e:
            logger.warning(f"Failed to replace field in single run: {e}")

    def _replace_field_across_runs_fixed(
        self, paragraph, affected_runs: List[Dict[str, Any]], replacement_text: str
    ) -> None:
        """FIXED: Replace field that spans across multiple runs."""
        try:
            if not affected_runs:
                return

            # FIXED STRATEGY: 
            # 1. Build the complete replacement text by combining all parts
            # 2. Put the entire replacement in the first run
            # 3. Clear the field portions from all other runs
            # 4. Preserve any text before the field in the first run
            # 5. Preserve any text after the field in the last run

            first_run_info = affected_runs[0]
            last_run_info = affected_runs[-1]
            
            # Get text before the field from the first run
            text_before_field = first_run_info["run_text"][:first_run_info["field_start_in_run"]]
            
            # Get text after the field from the last run
            text_after_field = last_run_info["run_text"][last_run_info["field_end_in_run"]:]
            
            # Build the complete new text for the first run
            new_first_run_text = text_before_field + replacement_text + text_after_field
            
            # Update the first run with the complete replacement
            first_run_info["run"].text = new_first_run_text
            
            # Clear all other affected runs
            for i in range(1, len(affected_runs)):
                affected_runs[i]["run"].text = ""
            
            logger.debug(f"Fixed multi-run replacement across {len(affected_runs)} runs")

        except Exception as e:
            logger.warning(f"Failed to replace field across runs (fixed): {e}")

    def _cleanup_empty_runs(self, paragraph) -> None:
        """Remove empty runs that can cause PowerPoint repair issues."""
        try:
            runs_to_remove = []

            for run in paragraph.runs:
                # Check if run is truly empty (no text or only whitespace)
                if not run.text or not run.text.strip():
                    # Check if run has meaningful formatting that should be preserved
                    has_formatting = False
                    if hasattr(run, "font"):
                        font = run.font
                        if (
                            font.bold is not None
                            or font.italic is not None
                            or font.underline is not None
                            or font.size is not None
                            or font.name is not None
                        ):
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
                    except Exception as remove_err:
                        logger.warning(f"Could not remove empty run: {remove_err}")

        except Exception as e:
            logger.warning(f"Failed to cleanup empty runs: {e}")

    def _get_field_value(self, field_name: str, data: Dict[str, Any]) -> Any:
        """Simple field value extraction for testing."""
        try:
            field_parts = field_name.split(".")
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
            
            return current_value if current_value is not None else ""
        except Exception as e:
            logger.warning(f"Failed to get field value for '{field_name}': {e}")
            return ""

# Key improvements in the fixed version:
print("""
KEY FIXES IMPLEMENTED:

1. **Process Fields in Reverse Order**: 
   - Fields are now processed from right to left (by field_start position)
   - This prevents earlier replacements from affecting later field positions

2. **Simplified Multi-Run Replacement Logic**:
   - Builds complete replacement text in one go
   - Places entire replacement in the first affected run
   - Clears all other affected runs completely
   - Preserves text before field (first run) and after field (last run)

3. **Eliminated Text Duplication**:
   - No longer tries to piece together text from multiple runs during replacement
   - Avoids the complex logic that was causing text corruption

4. **Cleaner State Management**:
   - Each run's final state is determined in one operation
   - No intermediate states that could cause inconsistencies

The original bug was in the multi-run replacement logic where it tried to
manage text across multiple runs simultaneously, leading to duplication and
corruption. The fix simplifies this by consolidating all replacement logic
into the first run and clearing the others.
""")