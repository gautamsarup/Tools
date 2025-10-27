# PowerPoint Text Extractor

A comprehensive Python tool for extracting text from PowerPoint presentations with OCR support, table extraction, and LLM-powered text formatting.

## Features

- **Native Text Extraction**: Extracts text from PowerPoint files using `python-pptx`
- **OCR Support**: Performs OCR on images embedded in slides for image-based content
- **Table Extraction**: Automatically extracts tables and exports them to Excel files
- **Selective Slide Processing**: Choose specific slides or process all slides
- **LLM Text Formatting**: Uses OpenAI API to format messy text into clean, readable content
- **Notes Extraction**: Extracts speaker notes from slides
- **Interactive Mode**: Command-line prompts for user-friendly operation

## Installation

### Prerequisites

1. **Python 3.7+** (recommended: Python 3.9+)
2. **Tesseract OCR** (for OCR functionality):
   - Windows: Download from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
   - macOS: `brew install tesseract`
   - Linux: `sudo apt-get install tesseract-ocr`

### Install Python Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Extract text from all slides:

```bash
python ppt_extractor.py presentation.pptx
```

### Extract Specific Slides

```bash
python ppt_extractor.py presentation.pptx --slides 1 3 5
```

### Interactive Mode

Let the program prompt you for slide selection:

```bash
python ppt_extractor.py presentation.pptx --interactive
```

### With OpenAI API Key

For LLM-powered text formatting:

```bash
python ppt_extractor.py presentation.pptx --openai-key YOUR_API_KEY
```

Or set it as an environment variable:

```bash
export OPENAI_API_KEY=your_api_key_here
python ppt_extractor.py presentation.pptx
```

### Custom Output Paths

```bash
python ppt_extractor.py presentation.pptx -o output.txt -e tables.xlsx
```

### Disable Features

Disable OCR:

```bash
python ppt_extractor.py presentation.pptx --no-ocr
```

Disable LLM formatting:

```bash
python ppt_extractor.py presentation.pptx --no-llm
```

### Tesseract Path (Windows)

If Tesseract is not auto-detected:

```bash
python ppt_extractor.py presentation.pptx --tesseract-path "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

## Command-Line Options

| Option | Description |
|--------|-------------|
| `pptx_path` | Path to PowerPoint file (required) |
| `-o, --output-text` | Output text file path |
| `-e, message` | Output Excel file path for tables |
| `--slides` | Specific slide numbers to extract (space-separated) |
| `--interactive` | Interactive mode for slide selection |
| `--no-ocr` | Disable OCR for images |
| `--no-llm` | Disable LLM text formatting |
| `--tesseract-path` | Path to tesseract executable |
| `--openai-key` | OpenAI API key |
| `-v, --verbose` | Enable verbose logging |

## Output Files

### Text Output (`_extracted.txt`)

Contains:
- Slide-by-slide text content
- Speaker notes
- OCR-extracted text from images
- Formatted text (if LLM enabled)

### Excel Output (`_tables.xlsx`)

Contains:
- Each table from the presentation as a separate sheet
- Sheet names: `Slide{N}_Table{M}`

## Examples

### Example 1: Extract All Content

```bash
python ppt_extractor.py quarterly_report.pptx
```

Output:
- `quarterly_report_extracted.txt` - All text content
- `quarterly_report_tables.xlsx` - All tables

### Example 2: Extract Specific Slides with LLM Formatting

```bash
python ppt_extractor.py quarterly_report.pptx --slides 1 5 10 --openai-key sk-...
```

### Example 3: Interactive Mode with Custom Output

```bash
python ppt_extractor.py presentation.pptx --interactive -o formatted_content.txt -e data_tables.xlsx
```

## How It Works

1. **Text Extraction**: Uses `python-pptx` to extract text from text frames, shapes, and grouped objects
2. **OCR Processing**: Converts images to text using Tesseract OCR
3. **Table Detection**: Identifies table shapes and extracts cell data
4. **Text Formatting**: Sends extracted text to OpenAI GPT-3.5-turbo for intelligent formatting
5. **Excel Export**: Converts table data to pandas DataFrames and exports to Excel

## OCR Configuration

The tool automatically detects Tesseract in common installation locations. If not found, you can:

1. Install Tesseract from the official repository
2. Specify the path using `--tesseract-path`
3. Ensure Tesseract is in your system PATH

## LLM Formatting

When enabled, the LLM formatting feature:
- Organizes scattered text into logical paragraphs
- Adds proper headings and sections
- Preserves lists and bullet points
- Improves overall readability

**Note**: Requires OpenAI API key and will make API calls to OpenAI.

## Troubleshooting

### "Missing required library: python-pptx 

Install dependencies:
```bash
pip install -r requirements.txt
```

### OCR Not Working

1. Verify Tesseract is installed:
   ```bash
   tesseract --version
   ```

2. Specify the path manually:
   ```bash
   python ppt_extractor.py file.pptx --tesseract-path /path/to/tesseract
   ```

### OpenAI API Errors

1. Check your API key is valid
2. Ensure you have credits/quota available
3. Try disabling LLM formatting: `--no-llm`

### No Tables Extracted

Tables must be actual PowerPoint table objects. Images of tables or text arranged to look like tables won't be extracted.

## Limitations

- Text in images requires OCR, which may have accuracy limitations
- Complex layouts may not preserve exact formatting
- Nested tables may not be handled correctly
- Very large presentations may take time to process

## Contributing

Feel free to submit issues or pull requests for improvements!

## License

This project is open source and available under the MIT License.

