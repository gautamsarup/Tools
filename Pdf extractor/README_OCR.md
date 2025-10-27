# PDF OCR Text Extractor

A comprehensive Python program that extracts text from PDF files using OCR, with special support for tables and multi-column layouts.

## Features

- **OCR Support**: Extracts text from scanned PDFs using Tesseract OCR
- **Table Extraction**: Preserves table formatting and structure
- **Multi-Column Handling**: Correctly parses multi-column documents
- **Multiple Extraction Methods**: Uses both pdfplumber and PyMuPDF for optimal results
- **Flexible Configuration**: Choose which features to enable/disable
- **Error Handling**: Robust error handling with detailed logging
- **Command Line Interface**: Easy-to-use CLI for batch processing

## Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements_ocr.txt
```

Or install individually:
```bash
pip install pdfplumber pytesseract Pillow PyMuPDF
```

### 2. Install Tesseract OCR

#### Windows
1. Download Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install the executable
3. Add Tesseract to your PATH or specify the path in the code

#### macOS
```bash
brew install tesseract
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get install tesseract-ocr
```

#### Linux (CentOS/RHEL)
```bash
sudo yum install tesseract
```

### 3. Optional: Enhanced Table Extraction

For better table extraction, install camelot-py:
```bash
pip install camelot-py[cv]
```

## Usage

### Command Line Interface

Basic usage:
```bash
python pdf_ocr_extractor.py document.pdf
```

With custom output file:
```bash
python pdf_ocr_extractor.py document.pdf -o extracted_text.txt
```

With custom tesseract path:
```bash
python pdf_ocr_extractor.py document.pdf --tesseract-path "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

Disable specific features:
```bash
# Disable OCR (for native PDFs only)
python pdf_ocr_extractor.py document.pdf --no-ocr

# Disable table extraction
python pdf_ocr_extractor.py document.pdf --no-tables

# Disable multi-column handling
python pdf_ocr_extractor.py document.pdf --no-multicolumn
```

Verbose output:
```bash
python pdf_ocr_extractor.py document.pdf --verbose
```

### Programmatic Usage

```python
from pdf_ocr_extractor import PDFOCRExtractor

# Initialize extractor
extractor = PDFOCRExtractor()

# Extract text with all features
result = extractor.extract_text_from_pdf(
    pdf_path="document.pdf",
    use_ocr=True,           # Enable OCR for scanned PDFs
    preserve_tables=True,   # Extract and format tables
    handle_multi_column=True # Handle multi-column layouts
)

# Print summary
print(f"Pages processed: {result['total_pages']}")
print(f"Tables found: {result['tables_found']}")

# Save extracted content
extractor.save_extracted_content(result, "output.txt")

# Access individual pages
for page in result['pages']:
    print(f"Page {page['page_number']}:")
    print(page['text'])
    
    # Format tables
    if page['tables']:
        formatted_tables = extractor.format_tables_as_text(page['tables'])
        print(formatted_tables)
```

### Custom Tesseract Path

```python
# Specify custom tesseract path
extractor = PDFOCRExtractor(tesseract_path=r"C:\Program Files\Tesseract-OCR\tesseract.exe")
```

## Output Format

The program generates a structured output containing:

- **Text Content**: Extracted text from each page
- **Tables**: Formatted tables with proper alignment
- **Metadata**: Extraction method, page count, table count
- **Errors**: Any errors encountered during processing

Example output structure:
```
PDF Text Extraction Results
File: document.pdf
Extraction Method: pdfplumber
Total Pages: 5
Tables Found: 3
==================================================

PAGE 1
------------------------------
TEXT CONTENT:
This is the extracted text from page 1...

TABLES:
--- Table 1 ---
Header 1    | Header 2    | Header 3
Value 1     | Value 2     | Value 3
Value 4     | Value 5     | Value 6

PAGE 2
------------------------------
...
```

## Supported PDF Types

### Native PDFs
- PDFs with selectable text
- Uses direct text extraction (fastest)
- Preserves original formatting

### Scanned PDFs
- Image-based PDFs
- Uses OCR to extract text
- Requires good image quality for best results

### Mixed PDFs
- Combination of native text and images
- Automatically detects and processes each type appropriately

## Table Extraction

The program handles various table formats:

- **Bordered Tables**: Tables with visible borders
- **Borderless Tables**: Tables without visible borders
- **Multi-page Tables**: Tables spanning multiple pages
- **Complex Tables**: Tables with merged cells and complex layouts

## Multi-Column Support

For documents with multiple columns:

- **Column Detection**: Automatically detects column boundaries
- **Reading Order**: Maintains proper left-to-right, top-to-bottom reading order
- **Layout Preservation**: Preserves the original document structure

## Error Handling

The program includes comprehensive error handling:

- **File Not Found**: Clear error messages for missing files
- **OCR Failures**: Graceful handling of OCR processing errors
- **Corrupted PDFs**: Robust handling of damaged or corrupted files
- **Permission Errors**: Clear messages for file access issues

## Performance Tips

1. **For Native PDFs**: Disable OCR for faster processing
2. **For Scanned PDFs**: Ensure good image quality
3. **For Large Files**: Process pages individually if memory is limited
4. **For Batch Processing**: Use the programmatic interface for automation

## Troubleshooting

### Common Issues

1. **Tesseract Not Found**
   - Install Tesseract OCR
   - Add to PATH or specify custom path
   - Check installation on your system

2. **Poor OCR Results**
   - Ensure PDF images are high quality
   - Try different Tesseract configurations
   - Consider preprocessing images

3. **Table Formatting Issues**
   - Some complex tables may not extract perfectly
   - Try different extraction methods
   - Consider manual review for critical tables

4. **Memory Issues with Large PDFs**
   - Process pages individually
   - Use lower resolution for OCR
   - Close other applications

### Getting Help

If you encounter issues:

1. Check the error messages in the output
2. Enable verbose logging with `--verbose`
3. Try with different PDF files to isolate the issue
4. Check that all dependencies are properly installed

## Examples

See `example_usage.py` for comprehensive examples including:

- Basic text extraction
- Table-focused extraction
- Batch processing
- Custom configuration
- Error handling

## License

This project is open source. Feel free to modify and distribute according to your needs.

## Contributing

Contributions are welcome! Areas for improvement:

- Enhanced table extraction algorithms
- Support for more OCR engines
- Better multi-column detection
- Performance optimizations
- Additional output formats (JSON, XML, etc.)
  
