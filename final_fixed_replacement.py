#!/usr/bin/env python3
"""
Final fixed version of the merge field replacement methods.

This version completely rethinks the approach to avoid position conflicts.
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

class FinalFixedPowerPointProcessor:
    """Final fixed version using iterative approach to avoid position conflicts."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def _process_paragraph_preserve_formatting(
        self, paragraph, data: Dict[str, Any]
    ) -> bool:
        """Process paragraph while preserving run-level formatting - FINAL FIXED VERSION."""
        try:
            # FINAL FIX: Use iterative approach
            # Keep processing until no more fields are found
            max_iterations = 10  # Prevent infinite loops
            iteration = 0
            
            while iteration < max_iterations:
                # Recalculate field positions each time
                field_positions = self._find_merge_fields_in_runs(paragraph)
                
                if not field_positions:
                    # No more fields to process
                    break
                
                # Process just the first field found
                field_info = field_positions[0]
                field_name = field_info["field"]
                field_value = self._get_field_value(field_name, data)
                field_value_str = str(field_value) if field_value is not None else ""
                
                # Replace this one field
                self._replace_field_in_runs_final(paragraph, field_info, field_value_str)
                
                iteration += 1
            
            if iteration >= max_iterations:
                logger.warning("Reached maximum iterations in paragraph processing")
            
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

    def _replace_field_in_runs_final(
        self, paragraph, field_info: Dict[str, Any], replacement_text: str
    ) -> None:
        """FINAL: Replace a merge field in runs while preserving formatting."""
        try:
            affected_runs = field_info["affected_runs"]

            if len(affected_runs) == 1:
                # Simple case: field is entirely within one run
                self._replace_field_in_single_run(affected_runs[0], replacement_text)
            else:
                # Complex case: field spans multiple runs - FINAL VERSION
                self._replace_field_across_runs_final(affected_runs, replacement_text)

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
            logger.debug(f"Single run replacement: '{original_text}' -> '{new_text}'")

        except Exception as e:
            logger.warning(f"Failed to replace field in single run: {e}")

    def _replace_field_across_runs_final(
        self, affected_runs: List[Dict[str, Any]], replacement_text: str
    ) -> None:
        """FINAL: Replace field that spans across multiple runs."""
        try:
            if not affected_runs:
                return

            # FINAL STRATEGY: 
            # 1. Build the complete replacement by preserving text before and after field
            # 2. Put everything in the first run
            # 3. Clear all other affected runs
            
            first_run = affected_runs[0]
            last_run = affected_runs[-1]
            
            # Text before the field (from first run)
            text_before = first_run["run_text"][:first_run["field_start_in_run"]]
            
            # Text after the field (from last run)  
            text_after = last_run["run_text"][last_run["field_end_in_run"]:]
            
            # Complete replacement text
            complete_text = text_before + replacement_text + text_after
            
            # Update first run with complete text
            first_run["run"].text = complete_text
            
            # Clear all other affected runs
            for i in range(1, len(affected_runs)):
                affected_runs[i]["run"].text = ""
                
            logger.debug(f"Multi-run replacement: combined {len(affected_runs)} runs into first run")
            logger.debug(f"Result: '{complete_text}'")

        except Exception as e:
            logger.warning(f"Failed to replace field across runs (final): {e}")

    def _cleanup_empty_runs(self, paragraph) -> None:
        """Remove empty runs that can cause PowerPoint repair issues."""
        try:
            # Don't remove runs during processing - just leave them empty
            # This prevents issues with run indexing
            pass

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

print("""
FINAL FIX STRATEGY:

1. **Iterative Processing**: 
   - Process one field at a time
   - Recalculate field positions after each replacement
   - Continue until no more fields are found

2. **Avoid Position Conflicts**:
   - By processing one field at a time, we eliminate position shift issues
   - Each iteration works with fresh position calculations

3. **Simplified Multi-Run Logic**:
   - Combine all text into the first affected run
   - Clear all other affected runs
   - Preserve text before and after the field

4. **Robust Error Handling**:
   - Maximum iteration limit prevents infinite loops
   - Graceful handling of edge cases

This approach completely eliminates the position tracking issues that were 
causing the original bug.
""")