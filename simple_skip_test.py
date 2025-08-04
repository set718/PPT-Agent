#!/usr/bin/env python
# -*- coding: utf-8 -*-

from dify_api_client import process_pages_with_dify

def simple_test():
    test_data = {
        "success": True,
        "pages": [
            {"page_number": 1, "page_type": "title", "title": "Title Page"},
            {"page_number": 2, "page_type": "content", "title": "Content Page 1"},
            {"page_number": 3, "page_type": "content", "title": "Content Page 2"},
            {"page_number": 4, "page_type": "ending", "title": "Ending Page"}
        ]
    }
    
    print("Testing fixed pages skip functionality...")
    print("Pages to test:")
    for page in test_data['pages']:
        print(f"  Page {page['page_number']}: {page['page_type']} - {page['title']}")
    
    result = process_pages_with_dify(test_data)
    
    print(f"\nResults:")
    summary = result.get('processing_summary', {})
    print(f"Total pages: {summary.get('total_pages', 0)}")
    print(f"Successful API calls: {summary.get('successful_api_calls', 0)}")
    print(f"Failed API calls: {summary.get('failed_api_calls', 0)}")
    print(f"Skipped fixed pages: {summary.get('skipped_fixed_pages', 0)}")
    
    print(f"\nExpected: Skip 2 pages (title + ending), Process 2 pages (content)")
    actual_skipped = summary.get('skipped_fixed_pages', 0)
    actual_processed = summary.get('successful_api_calls', 0) + summary.get('failed_api_calls', 0)
    
    if actual_skipped == 2 and actual_processed == 2:
        print("SUCCESS: Fixed pages skip working correctly!")
    else:
        print(f"ERROR: Expected skip=2, process=2. Got skip={actual_skipped}, process={actual_processed}")
    
    return result

if __name__ == "__main__":
    simple_test()