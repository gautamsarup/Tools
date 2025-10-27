# Quick Start Guide

## Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

## Step 2: Set Up Your OpenAI API Key

You have three options:

### Option A: Use the Setup Script (Recommended)

```bash
python setup_api_key.py
```

This will prompt you for your API key and save it securely to `config.py`.

### Option B: Manual Config File

1. Copy the template:
   ```bash
   cp config_template.py config.py
   ```

2. Edit `config.py` and add your API key:
   ```python
   OPENAI_API_KEY = "sk-your-actual-api-key-here"
   ```

### Option C: Environment Variable

```bash
# Windows PowerShell
$env:OPENAI_API_KEY="sk-your-api-key-here"

# Windows CMD
set OPENAI_API_KEY=sk-your-api-key-here

# Linux/macOS
export OPENAI_API_KEY="sk-your-api-key-here"
```

## Step 3: Run the Extractor

### Basic Usage

```bash
python ppt_extractor.py your_presentation.pptx
```

### Interactive Mode (Choose Specific Slides)

```bash
python ppt_extractor.py your_presentation.pptx --interactive
```

### Extract Specific Slides

```bash
python ppt_extractor.py your_presentation.pptx --slides 1 3 5
```

## Output Files

After running, you'll get:
- `{filename}_extracted.txt` - Formatted text content
- `{filename}_tables.xlsx` - All tables in Excel format

## What If I Don't Have an OpenAI API Key?

You can still use the extractor without LLM formatting:

```bash
python ppt_extractor.py your_presentation.pptx --no-llm
```

You'll still get:
- ✅ Text extraction from slides
- ✅ OCR from images
- ✅ Table extraction to Excel
- ❌ LLM-powered text formatting (disabled)

## Troubleshooting

### "Module not found" error
Run: `pip install -r requirements.txt`

### OCR not working
Install Tesseract OCR:
- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
- macOS: `brew install tesseract`
- Linux: `sudo apt-get install tesseract-ocr`

### API key issues
Make sure your key starts with `sk-` and is valid.

## Need Help?

Check the full [README.md](README.md) for detailed documentation.

