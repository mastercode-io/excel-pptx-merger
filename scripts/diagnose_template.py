#!/usr/bin/env python3
"""PowerPoint template merge field diagnostic script."""

import json
import sys
import argparse
from pathlib import Path
from typing import Dict, List

# Add src to path
script_dir = Path(__file__).parent
project_root = script_dir.parent
sys.path.insert(0, str(project_root / 'src'))

from src.pptx_processor import PowerPointProcessor
from src.utils.exceptions import PowerPointProcessingError


def diagnose_template(template_path: str) -> Dict[str, List[str]]:
    """Diagnose PowerPoint template and return merge fields detected per slide.
    
    Args:
        template_path: Path to the PowerPoint template file
        
    Returns:
        Dictionary with slide numbers as keys and lists of detected merge fields as values
    """
    try:
        processor = PowerPointProcessor(template_path)
        
        # Get all merge fields from the template
        all_fields = processor.get_merge_fields()
        
        # Get slide-specific merge fields
        slide_fields = {}
        
        if processor.presentation:
            for slide_idx, slide in enumerate(processor.presentation.slides):
                slide_number = slide_idx + 1
                fields = processor._extract_slide_merge_fields(slide)
                slide_fields[f"slide_{slide_number}"] = sorted(fields)
        
        return {
            "template_path": template_path,
            "total_slides": len(processor.presentation.slides) if processor.presentation else 0,
            "all_merge_fields": sorted(all_fields),
            "slide_fields": slide_fields
        }
        
    except PowerPointProcessingError as e:
        return {
            "error": f"PowerPoint processing error: {str(e)}",
            "template_path": template_path
        }
    except Exception as e:
        return {
            "error": f"Unexpected error: {str(e)}",
            "template_path": template_path
        }


def main():
    """Main function for the diagnostic script."""
    parser = argparse.ArgumentParser(description="Diagnose PowerPoint template merge fields")
    parser.add_argument("template", help="Path to PowerPoint template file")
    parser.add_argument("-o", "--output", help="Output file path (default: print to stdout)")
    parser.add_argument("--pretty", action="store_true", help="Pretty print JSON output")
    
    args = parser.parse_args()
    
    # Validate template file exists
    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Error: Template file not found: {template_path}", file=sys.stderr)
        sys.exit(1)
    
    # Diagnose template
    print(f"Diagnosing PowerPoint template: {template_path}")
    results = diagnose_template(str(template_path))
    
    # Format output
    if args.pretty:
        json_output = json.dumps(results, indent=2, ensure_ascii=False)
    else:
        json_output = json.dumps(results, ensure_ascii=False)
    
    # Write output
    if args.output:
        output_path = Path(args.output)
        output_path.write_text(json_output, encoding='utf-8')
        print(f"Results written to: {output_path}")
    else:
        print("\nDiagnostic Results:")
        print(json_output)
    
    # Summary
    if "error" not in results:
        total_slides = results.get("total_slides", 0)
        total_fields = len(results.get("all_merge_fields", []))
        print(f"\nSummary: {total_slides} slides, {total_fields} unique merge fields detected")
        
        # Show fields per slide
        for slide_key, fields in results.get("slide_fields", {}).items():
            print(f"  {slide_key}: {len(fields)} fields")
    else:
        print(f"\nError: {results['error']}")
        sys.exit(1)


if __name__ == "__main__":
    main()