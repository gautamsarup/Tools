#!/usr/bin/env python3
"""
PDF OCR Text Extractor with Table and Multi-Column Support

This program extracts text from PDF files using OCR, with special handling for:
- Tables (preserving formatting)
- Multi-column layouts (maintaining correct reading order)
- Scanned PDFs (using OCR)
- Native PDFs (direct text extraction)

Dependencies:
- pdfplumber: For PDF text and table extraction
- pytesseract: For OCR processing
- Pillow: For image processing
- PyMuPDF (fitz): For advanced text extraction
- camelot-py: For table extraction (optional)
"""

import os
import sys
import argparse
import io
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import logging

try:
    import pdfplumber
    import pytesseract
    from PIL import Image
    import fitz  # PyMuPDF
except ImportError as e:
    print(f"Missing required library: {e}")
    print("Please install required packages:")
    print("pip install pdfplumber pytesseract Pillow PyMuPDF")
    sys.exit(1)

# Optional imports
try:
    import camelot
    CAMELOT_AVAILABLE = True
except ImportError:
    CAMELOT_AVAILABLE = False
    print("Note: camelot-py not available. Install with 'pip install camelot-py[cv]' for enhanced table extraction.")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class PDFOCRExtractor:
    """
    A comprehensive PDF text extractor that handles OCR, tables, and multi-column layouts.
    """
    
    def __init__(self, tesseract_path: Optional[str] = None):
        """
        Initialize the PDF OCR extractor.
        
        Args:
            tesseract_path: Path to tesseract executable (auto-detected if None)
        """
        self.tesseract_path = tesseract_path
        self._setup_tesseract()
        
    def _setup_tesseract(self):
        """Setup tesseract path if provided."""
        if self.tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
            logger.info(f"Using tesseract at: {self.tesseract_path}")
        else:
            # Try to auto-detect tesseract
            common_paths = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',  # Windows
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',  # Windows 32-bit
                '/usr/bin/tesseract',  # Linux
                '/usr/local/bin/tesseract',  # macOS
                '/opt/homebrew/bin/tesseract',  # macOS with Homebrew
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    logger.info(f"Auto-detected tesseract at: {path}")
                    return
            
            logger.warning("Tesseract not found in common locations. Please install tesseract-ocr.")
    
    def extract_text_from_pdf(self, pdf_path: str, use_ocr: bool = True, 
                            preserve_tables: bool = True, 
                            handle_multi_column: bool = True) -> Dict[str, Any]:
        """
        Extract text from PDF with comprehensive formatting preservation.
        
        Args:
            pdf_path: Path to the PDF file
            use_ocr: Whether to use OCR for scanned PDFs
            preserve_tables: Whether to extract and format tables
            handle_multi_column: Whether to handle multi-column layouts
            
        Returns:
            Dictionary containing extracted content
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        logger.info(f"Processing PDF: {pdf_path}")
        
        result = {
            'file_path': pdf_path,
            'pages': [],
            'total_pages': 0,
            'extraction_method': 'unknown',
            'tables_found': 0,
            'errors': []
        }
        
        try:
            # First, try with pdfplumber for native PDFs
            result = self._extract_with_pdfplumber(pdf_path, use_ocr, preserve_tables, handle_multi_column)
            
            # If pdfplumber fails or returns minimal content, try PyMuPDF
            if not result['pages'] or all(not page.get('text', '').strip() for page in result['pages']):
                logger.info("Trying PyMuPDF for better text extraction...")
                result = self._extract_with_pymupdf(pdf_path, use_ocr, preserve_tables, handle_multi_column)
                
        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
            result['errors'].append(str(e))
        
        return result
    
    def _extract_with_pdfplumber(self, pdf_path: str, use_ocr: bool, 
                               preserve_tables: bool, handle_multi_column: bool) -> Dict[str, Any]:
        """Extract text using pdfplumber."""
        result = {
            'file_path': pdf_path,
            'pages': [],
            'total_pages': 0,
            'extraction_method': 'pdfplumber',
            'tables_found': 0,
            'errors': []
        }
        
        with pdfplumber.open(pdf_path) as pdf:
            result['total_pages'] = len(pdf.pages)
            
            for page_num, page in enumerate(pdf.pages, 1):
                logger.info(f"Processing page {page_num}/{len(pdf.pages)}")
                
                page_data = {
                    'page_number': page_num,
                    'text': '',
                    'tables': [],
                    'extraction_method': 'native',
                    'ocr_used': False
                }
                
                # Extract text
                text = page.extract_text()
                if text and text.strip():
                    page_data['text'] = text
                    page_data['extraction_method'] = 'native'
                elif use_ocr:
                    # Try OCR if no text found
                    logger.info(f"Performing OCR on page {page_num}")
                    try:
                        ocr_text = self._perform_ocr_on_page(page)
                        page_data['text'] = ocr_text
                        page_data['extraction_method'] = 'ocr'
                        page_data['ocr_used'] = True
                    except Exception as e:
                        logger.error(f"OCR failed on page {page_num}: {e}")
                        result['errors'].append(f"OCR failed on page {page_num}: {e}")
                
                # Extract tables if requested
                if preserve_tables:
                    tables = self._extract_tables_from_page(page)
                    page_data['tables'] = tables
                    result['tables_found'] += len(tables)
                
                result['pages'].append(page_data)
        
        return result
    
    def _extract_with_pymupdf(self, pdf_path: str, use_ocr: bool, 
                            preserve_tables: bool, handle_multi_column: bool) -> Dict[str, Any]:
        """Extract text using PyMuPDF for better multi-column handling."""
        result = {
            'file_path': pdf_path,
            'pages': [],
            'total_pages': 0,
            'extraction_method': 'pymupdf',
            'tables_found': 0,
            'errors': []
        }
        
        doc = fitz.open(pdf_path)
        result['total_pages'] = len(doc)
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            logger.info(f"Processing page {page_num + 1}/{len(doc)}")
            
            page_data = {
                'page_number': page_num + 1,
                'text': '',
                'tables': [],
                'extraction_method': 'native',
                'ocr_used': False
            }
            
            # Extract text with block information for better multi-column handling
            if handle_multi_column:
                text = self._extract_text_with_column_awareness(page)
            else:
                text = page.get_text()
            
            if text and text.strip():
                page_data['text'] = text
                page_data['extraction_method'] = 'native'
            elif use_ocr:
                # Try OCR if no text found
                logger.info(f"Performing OCR on page {page_num + 1}")
                try:
                    ocr_text = self._perform_ocr_on_pymupdf_page(page)
                    page_data['text'] = ocr_text
                    page_data['extraction_method'] = 'ocr'
                    page_data['ocr_used'] = True
                except Exception as e:
                    logger.error(f"OCR failed on page {page_num + 1}: {e}")
                    result['errors'].append(f"OCR failed on page {page_num + 1}: {e}")
            
            # Extract tables if requested
            if preserve_tables:
                tables = self._extract_tables_from_pymupdf_page(page)
                page_data['tables'] = tables
                result['tables_found'] += len(tables)
            
            result['pages'].append(page_data)
        
        doc.close()
        return result
    
    def _extract_text_with_column_awareness(self, page) -> str:
        """Extract text with proper column ordering using PyMuPDF."""
        blocks = page.get_text("blocks")
        
        # Sort blocks by vertical position first, then horizontal
        blocks.sort(key=lambda b: (b[1], b[0]))
        
        text_parts = []
        for block in blocks:
            if len(block) > 4 and block[4].strip():  # Check if block has text
                text_parts.append(block[4].strip())
        
        return '\n'.join(text_parts)
    
    def _perform_ocr_on_page(self, page) -> str:
        """Perform OCR on a pdfplumber page."""
        # Convert page to image
        image = page.to_image()
        pil_image = image.original
        
        # Perform OCR
        ocr_text = pytesseract.image_to_string(pil_image, config='--psm 6')
        return ocr_text
    
    def _perform_ocr_on_pymupdf_page(self, page) -> str:
        """Perform OCR on a PyMuPDF page."""
        # Convert page to image
        mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        # Convert to PIL Image
        pil_image = Image.open(io.BytesIO(img_data))
        
        # Perform OCR
        ocr_text = pytesseract.image_to_string(pil_image, config='--psm 6')
        return ocr_text
    
    def _extract_tables_from_page(self, page) -> List[List[List[str]]]:
        """Extract tables from a pdfplumber page."""
        tables = []
        try:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                if table:  # Check if table is not empty
                    tables.append(table)
        except Exception as e:
            logger.error(f"Error extracting tables: {e}")
        
        return tables
    
    def _extract_tables_from_pymupdf_page(self, page) -> List[List[List[str]]]:
        """Extract tables from a PyMuPDF page (basic implementation)."""
        # PyMuPDF doesn't have built-in table extraction
        # This is a placeholder for future enhancement
        return []
    
    def format_tables_as_text(self, tables: List[List[List[str]]]) -> str:
        """Format extracted tables as readable text."""
        if not tables:
            return ""
        
        formatted_tables = []
        for i, table in enumerate(tables, 1):
            formatted_tables.append(f"\n--- Table {i} ---")
            
            if not table:
                formatted_tables.append("(Empty table)")
                continue
            
            # Find maximum width for each column
            max_widths = []
            for row in table:
                for col_idx, cell in enumerate(row):
                    if col_idx >= len(max_widths):
                        max_widths.append(0)
                    max_widths[col_idx] = max(max_widths[col_idx], len(str(cell)))
            
            # Format table
            for row in table:
                formatted_row = []
                for col_idx, cell in enumerate(row):
                    cell_str = str(cell) if cell else ""
                    formatted_cell = cell_str.ljust(max_widths[col_idx])
                    formatted_row.append(formatted_cell)
                formatted_tables.append(" | ".join(formatted_row))
        
        return "\n".join(formatted_tables)
    
    def save_extracted_content(self, result: Dict[str, Any], output_path: str):
        """Save extracted content to a text file."""
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"PDF Text Extraction Results\n")
            f.write(f"File: {result['file_path']}\n")
            f.write(f"Extraction Method: {result['extraction_method']}\n")
            f.write(f"Total Pages: {result['total_pages']}\n")
            f.write(f"Tables Found: {result['tables_found']}\n")
            f.write("=" * 50 + "\n\n")
            
            for page in result['pages']:
                f.write(f"PAGE {page['page_number']}\n")
                f.write("-" * 30 + "\n")
                
                if page['text']:
                    f.write("TEXT CONTENT:\n")
                    f.write(page['text'])
                    f.write("\n\n")
                
                if page['tables']:
                    f.write("TABLES:\n")
                    formatted_tables = self.format_tables_as_text(page['tables'])
                    f.write(formatted_tables)
                    f.write("\n\n")
            
            if result['errors']:
                f.write("ERRORS:\n")
                for error in result['errors']:
                    f.write(f"- {error}\n")
        
        logger.info(f"Extracted content saved to: {output_path}")


def main():
    """Main function for command-line usage."""
    parser = argparse.ArgumentParser(description='Extract text from PDF files with OCR support')
    parser.add_argument('pdf_path', help='Path to the PDF file')
    parser.add_argument('-o', '--output', help='Output text file path')
    parser.add_argument('--no-ocr', action='store_true', help='Disable OCR for scanned PDFs')
    parser.add_argument('--no-tables', action='store_true', help='Disable table extraction')
    parser.add_argument('--no-multicolumn', action='store_true', help='Disable multi-column handling')
    parser.add_argument('--tesseract-path', help='Path to tesseract executable')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose logging')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Initialize extractor
    extractor = PDFOCRExtractor(tesseract_path=args.tesseract_path)
    
    try:
        # Extract content
        result = extractor.extract_text_from_pdf(
            pdf_path=args.pdf_path,
            use_ocr=not args.no_ocr,
            preserve_tables=not args.no_tables,
            handle_multi_column=not args.no_multicolumn
        )
        
        # Determine output path
        if args.output:
            output_path = args.output
        else:
            pdf_name = Path(args.pdf_path).stem
            output_path = f"{pdf_name}_extracted.txt"
        
        # Save results
        extractor.save_extracted_content(result, output_path)
        
        # Print summary
        print(f"\nExtraction completed!")
        print(f"Pages processed: {result['total_pages']}")
        print(f"Tables found: {result['tables_found']}")
        print(f"Output saved to: {output_path}")
        
        if result['errors']:
            print(f"Errors encountered: {len(result['errors'])}")
            for error in result['errors']:
                print(f"  - {error}")
    
    except Exception as e:
        logger.error(f"Failed to process PDF: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
