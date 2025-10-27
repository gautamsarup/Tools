#!/usr/bin/env python3
"""
Setup script for configuring OpenAI API key securely.
"""

import os
import sys
from pathlib import Path


def setup_api_key():
    """Interactive setup for OpenAI API key."""
    print("=" * 60)
    print("PowerPoint Extractor - API Key Setup")
    print("=" * 60)
    print()
    
    # Check if config.py already exists
    config_file = Path("config.py")
    if config_file.exists():
        print("⚠️  config.py already exists!")
        response = input("Do you want to overwrite it? (yes/no): ").strip().lower()
        if response != 'yes':
            print("Setup cancelled.")
            return
    
    # Get API key from user
    print("Please enter your OpenAI API key:")
    print("(It will be saved locally and NOT committed to git)")
    print()
    
    api_key = input("OpenAI API Key: ").strip()
    
    if not api_key:
        print("❌ No API key provided. Setup cancelled.")
        return
    
    if not api_key.startswith('sk-'):
        print("⚠️  Warning: OpenAI API keys typically start with 'sk-'")
        response = input("Continue anyway? (yes/no): ").strip().lower()
        if response != 'yes':
            print("Setup cancelled.")
            return
    
    # Get optional Tesseract path
    print()
    print("Optional: Enter Tesseract path (or press Enter to skip auto-detection):")
    tesseract_path = input("Tesseract path: ").strip()
    
    # Create config.py
    config_content = f'''#!/usr/bin/env python3
"""
Configuration for PowerPoint Extractor

This file contains your API keys and settings.
DO NOT commit this file to version control!
"""

# OpenAI API Configuration
OPENAI_API_KEY = "{api_key}"

# Tesseract OCR Configuration (optional - leave as None for auto-detection)
TESSERACT_PATH = {repr(tesseract_path) if tesseract_path else 'None'}

# Default Settings
DEFAULT_USE_OCR = True
DEFAULT_USE_LLM_FORMATTING = True
'''
    
    try:
        with open('config.py', 'w') as f:
            f.write(config_content)
        
        print()
        print("✅ Configuration saved successfully!")
        print()
        print("📝 Your API key has been saved to config.py")
        print("🔒 This file is excluded from git (already in .gitignore)")
        print()
        print("Now you can use the extractor without specifying --openai-key:")
        print("  python ppt_extractor.py your_presentation.pptx")
        print()
        
    except Exception as e:
        print(f"❌ Error saving configuration: {e}")
        sys.exit(1)


def setup_environment_variable():
    """Alternative: Set up environment variable."""
    print()
    print("=" * 60)
    print("Alternative: Using Environment Variable")
    print("=" * 60)
    print()
    print("You can also set your API key as an environment variable:")
    print()
    print("Windows (PowerShell):")
    print('  $env:OPENAI_API_KEY="your-api-key-here"')
    print()
    print("Windows (Command Prompt):")
    print('  set OPENAI_API_KEY=your-api-key-here')
    print()
    print("Linux/macOS:")
    print('  export OPENAI_API_KEY="your-api-key-here"')
    print()
    print("To make it permanent (Linux/macOS), add to ~/.bashrc or ~/.zshrc:")
    print('  echo \'export OPENAI_API_KEY="your-api-key-here"\' >> ~/.bashrc')
    print()


if __name__ == "__main__":
    setup_api_key()
    setup_environment_variable()

