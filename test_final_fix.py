#!/usr/bin/env python3
"""
Test script to verify the final fixed merge field replacement works correctly.
"""

import logging
from final_fixed_replacement import FinalFixedPowerPointProcessor
from debug_merge_fields import MockParagraph

# Set up logging
logging.basicConfig(level=logging.DEBUG)

def test_final_fix():
    """Test the final fixed replacement logic against problematic scenarios."""
    
    print("TESTING FINAL FIXED REPLACEMENT LOGIC")
    print("=" * 60)
    
    # Test data
    test_data = {
        "word_marks": [
            {
                "application_number": "12345678",
                "mark_text": "SAMPLE MARK",
                "registration_number": "87654321",
            }
        ]
    }
    
    processor = FinalFixedPowerPointProcessor()
    
    # Test scenarios that were failing
    scenarios = [
        {
            "name": "Split Across Runs - Previously Broken",
            "runs": ['Application Number: {{word_marks.0.application_', 'number}}, Mark Text: {{word_marks.0.mark_text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        },
        {
            "name": "Different Split Pattern",
            "runs": ['Application Number: {{word_marks.0.', 'application_number}}, Mark Text: {{word_marks.0.mark_text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        },
        {
            "name": "Multiple Splits for One Field",
            "runs": ['Application Number: {{word_marks.', '0.', 'application_number}}, Mark Text: {{word_marks.0.mark_text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        },
        {
            "name": "Both Fields Split",
            "runs": ['Application Number: {{word_marks.0.application_', 'number}}, Mark Text: {{word_marks.0.mark_', 'text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        },
        {
            "name": "Field Pattern Split at Braces",
            "runs": ['Application Number: {', '{word_marks.0.application_number}', '}, Mark Text: {{word_marks.0.mark_text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        },
        {
            "name": "Single Run - Control Test",
            "runs": ['Application Number: {{word_marks.0.application_number}}, Mark Text: {{word_marks.0.mark_text}}'],
            "expected": "Application Number: 12345678, Mark Text: SAMPLE MARK"
        }
    ]
    
    results = []
    
    for scenario in scenarios:
        print(f"\n{'='*50}")
        print(f"TESTING: {scenario['name']}")
        print(f"{'='*50}")
        
        # Create mock paragraph
        paragraph = MockParagraph(scenario['runs'])
        
        print(f"Initial runs: {[run.text for run in paragraph.runs]}")
        
        try:
            # Apply the fixed processing
            success = processor._process_paragraph_preserve_formatting(paragraph, test_data)
            
            if success:
                # Check the result
                final_text = "".join(run.text for run in paragraph.runs if run.text)
                print(f"Final runs: {[run.text for run in paragraph.runs]}")
                print(f"Final text: '{final_text}'")
                print(f"Expected:   '{scenario['expected']}'")
                
                if final_text == scenario['expected']:
                    print("✅ SUCCESS: Text matches expected result")
                    results.append({"scenario": scenario['name'], "status": "PASS", "result": final_text})
                else:
                    print("❌ FAILURE: Text does not match expected result")
                    results.append({"scenario": scenario['name'], "status": "FAIL", "result": final_text, "expected": scenario['expected']})
            else:
                print("❌ FAILURE: Processing returned False")
                results.append({"scenario": scenario['name'], "status": "ERROR", "result": "Processing failed"})
                
        except Exception as e:
            print(f"❌ ERROR: {e}")
            import traceback
            traceback.print_exc()
            results.append({"scenario": scenario['name'], "status": "ERROR", "result": str(e)})
    
    # Summary
    print(f"\n{'='*60}")
    print("TEST RESULTS SUMMARY")
    print(f"{'='*60}")
    
    passed = sum(1 for r in results if r['status'] == 'PASS')
    failed = sum(1 for r in results if r['status'] == 'FAIL')
    errors = sum(1 for r in results if r['status'] == 'ERROR')
    
    print(f"Total scenarios: {len(results)}")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    print(f"Errors: {errors}")
    
    for result in results:
        status_icon = "✅" if result['status'] == 'PASS' else "❌"
        print(f"{status_icon} {result['scenario']}: {result['status']}")
        if result['status'] != 'PASS':
            print(f"   Result: {result['result']}")
            if 'expected' in result:
                print(f"   Expected: {result['expected']}")
    
    return passed == len(results)

if __name__ == "__main__":
    all_passed = test_final_fix()
    if all_passed:
        print(f"\n🎉 ALL TESTS PASSED! The final fix works correctly.")
    else:
        print(f"\n⚠️  Some tests failed. Review the results above.")