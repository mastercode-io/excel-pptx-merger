#!/usr/bin/env python3
"""
Test the fix in the actual production code.
"""

import sys
import os
sys.path.insert(0, os.path.abspath('/Users/alexsherin/Projects_/excel-pptx-merger'))

from src.pptx_processor import PowerPointProcessor
from debug_merge_fields import MockParagraph
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)

def test_production_fix():
    """Test that the production code fix works."""
    
    print("TESTING FIX IN PRODUCTION CODE")
    print("=" * 50)
    
    # Test data
    test_data = {
        "word_marks": [
            {
                "application_number": "12345678",
                "mark_text": "SAMPLE MARK",
            }
        ]
    }
    
    # Create a minimal PowerPoint processor (we'll only test the paragraph method)
    class TestProcessor(PowerPointProcessor):
        def __init__(self):
            # Skip the file validation
            self.template_path = "dummy"
            self.presentation = None
    
    processor = TestProcessor()
    
    # Test problematic scenario
    runs = ['Application Number: {{word_marks.0.application_', 'number}}, Mark Text: {{word_marks.0.mark_text}}']
    paragraph = MockParagraph(runs)
    
    print(f"Initial runs: {[run.text for run in paragraph.runs]}")
    
    try:
        # Test the fixed method
        success = processor._process_paragraph_preserve_formatting(paragraph, test_data)
        
        if success:
            final_text = "".join(run.text for run in paragraph.runs if run.text)
            print(f"Final runs: {[run.text for run in paragraph.runs]}")
            print(f"Final text: '{final_text}'")
            
            expected = "Application Number: 12345678, Mark Text: SAMPLE MARK"
            if final_text == expected:
                print("✅ SUCCESS: Production fix works correctly!")
                return True
            else:
                print(f"❌ FAILURE: Expected '{expected}'")
                return False
        else:
            print("❌ FAILURE: Processing returned False")
            return False
            
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_production_fix()
    if success:
        print(f"\n🎉 PRODUCTION FIX VERIFIED!")
    else:
        print(f"\n⚠️  Production fix failed.")