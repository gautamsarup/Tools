#!/usr/bin/env python3
"""
Example usage of the PowerPoint Extractor

This script demonstrates how to use the PowerPointExtractor class programmatically.
"""

from ppt_extractor import PowerPointExtractor
import os


def example_basic_extraction():
    """Example: Basic text extraction from all slides."""
    print("Example 1: Basic extraction from all slides")
    print("-" * 50)
    
    extractor = PowerPointExtractor()
    
    # Extract from all slides
    result = extractor.extract_text_from_pptx(
        pptx_path="sample_presentation.pptx",
        use_ocr=True,
        use_llm_formatting=False
    )
    
    # Save results
    extractor.save_text_content(result, "output_all_slides.txt")
    
    if result['tables_found'] > 0:
        extractor.save_tables_to_excel(result, "output_all_tables.xlsx")
    
    print(f"Processed {result['slides_processed']} slides")
    print(f"Found {result['tables_found']} tables")


def example_specific_slides():
    """Example: Extract from specific slides."""
    print("\nExample 2: Extract from specific slides")
    print("-" * 50)
    
    extractor = PowerPointExtractor()
    
    # Extract from slides 1, 3, and 5
    result = extractor.extract_text_from_pptx(
        pptx_path="sample_presentation.pptx",
        slide_numbers=[1, 3, 5],
        use_ocr=True,
        use_llm_formatting=False
    )
    
    extractor.save_text_content(result, "output_specific_slides.txt")
    
    print(f"Processed slides: {result['slides_processed']}")


def example_with_llm_formatting():
    """Example: Use LLM for text formatting."""
    print("\nExample 3: Extract with LLM formatting")
    print("-" * 50)
    
    # Get API key from environment or user input
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        print("Warning: OPENAI_API_KEY not set. LLM formatting will be disabled.")
        api_key = None
    
    extractor = PowerPointExtractor(openai_api_key=api_key)
    
    result = extractor.extract_text_from_pptx(
        pptx_path="sample_presentation.pptx",
        use_ocr=True,
        use_llm_formatting=True
    )
    
    extractor.save_text_content(result, "output_llm_formatted.txt")
    
    print(f"LLM formatting used: {result['slides'][0].get('llm_formatted', False)}")


def example_custom_output_paths():
    """Example: Custom output paths."""
    print("\nExample 4: Custom output paths")
    print("-" * 50)
    
    extractor = PowerPointExtractor()
    
    result = extractor.extract_text_from_pptx(
        pptx_path="sample_presentation.pptx",
        slide_numbers=[1, 2],
        use_ocr=False,
        use_llm_formatting=False
    )
    
    # Custom output paths
    extractor.save_text_content(result, "custom_text_output.txt")
    
    if result['tables_found'] > 0:
        extractor.save_tables_to_excel(result, "custom_tables_output.xlsx")
    
    print("Output saved to custom paths")


def example_access_extracted_data():
    """Example: Access and manipulate extracted data."""
    print("\nExample 5: Access extracted data programmatically")
    print("-" * 50)
    
    extractor = PowerPointExtractor()
    
    result = extractor.extract_text_from_pptx(
        pptx_path="sample_presentation.pptx",
        use_ocr=True,
        use_llm_formatting=False
    )
    
    # Access data from each slide
    for slide in result['slides']:
        print(f"\nSlide {slide['slide_number']}:")
        print(f"  Text length: {len(slide['text'])} characters")
        print(f"  Tables found: {len(slide['tables'])}")
        print(f"  OCR used: {slide['ocr_used']}")
        
        # Access table data
        for i, table in enumerate(slide['tables'], 1):
            print(f"  Table {i}: {len(table)} rows x {len(table[0]) if table else 0} columns")


if __name__ == "__main__":
    print("PowerPoint Extractor - Example Usage")
    print("=" * 50)
    
    # Replace with your actual PowerPoint file
    test_file = "sample_presentation.pptx"
    
    if not os.path.exists(test_file):
        print(f"\nError: {test_file} not found.")
        print("Please provide a PowerPoint file path.")
        print("\nTo run these examples:")
        print("1. Update the 'pptx_path' in each function")
        print("2. Run: python example_usage.py")
    else:
        # Run examples
        example_basic_extraction()
        example_specific_slides()
        example_with_llm_formatting()
        example_custom_output_paths()
        example_access_extracted_data()
        
        print("\n" + "=" * 50)
        print("Examples completed!")

