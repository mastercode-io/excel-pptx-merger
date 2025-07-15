# Range Image Debugging Guide

## Overview

The enhanced range image debugging system provides clear, visually distinct logging for range image extraction operations. This solves the issue of verbose PowerPoint processor logs drowning out important range image processing information.

## Quick Start

### 1. Enable Range Image Debug Mode

#### CLI Merge Command
```bash
# Enable enhanced range image debugging for merge operations
uv run python -m src.main merge -e excel_file.xlsx -p template.pptx --debug-range-images

# Combine with regular debug if needed
uv run python -m src.main merge -e excel_file.xlsx -p template.pptx --debug-images --debug-range-images
```

#### Server Mode
```bash
# Start server with range image debugging
uv run python -m src.main serve --debug-range-images

# Combine with regular debug mode
uv run python -m src.main serve --debug --debug-range-images
```

### 2. What You'll See

When range image debug mode is enabled, you'll see clearly formatted debug output like this:

```
================================================================================
🖼️  RANGE IMAGE DEBUG: INFO
📍 range_images - info:39
💬 🚀 EXTRACTION STARTED with 4 configurations
================================================================================

================================================================================
🖼️  RANGE IMAGE DEBUG: INFO
📍 range_images - info:39
💬 📋 RANGE CONFIG [0]:
   🏷️  Field Name: client_summary_table
   📊 Sheet Name: Order Form
   📍 Range: A1:F3
   🎨 Format: png
   📐 DPI: 150
   📏 Dimensions: auto x auto
================================================================================
```

## Debug Output Features

### 🎯 Visual Separation
- **Bold borders**: Each debug message is clearly separated with `=` borders
- **Emoji indicators**: Quick visual identification of message types
- **Structured format**: Consistent formatting across all range image logs

### 📊 Detailed Information
- **Configuration validation**: See exactly which configs are loaded and validated
- **Sheet availability**: View available vs requested sheet names
- **Progress tracking**: Visual progress bars for multi-range exports
- **Graph API status**: Real-time authentication and connection status
- **Export results**: Detailed success/failure information with file paths and dimensions

### 🔇 Noise Reduction
When `--debug-range-images` is enabled, the following loggers are automatically set to WARNING level:
- `src.pptx_processor` (reduces PowerPoint processing verbosity)
- `PIL` (reduces image processing noise)
- `matplotlib` (reduces plotting library noise)

## Configuration Testing

### Test Your Configuration
Use the included test script to validate your range image configuration:

```bash
python test_range_debug.py
```

This will:
- Validate your `config/range_images_example_config.json`
- Show detailed configuration information
- Simulate the debug output you'll see during actual processing

### Sample Configuration
The system includes a sample configuration at `config/range_images_example_config.json` with 4 range image examples.

## Troubleshooting Your Range Images

### Common Issues and Debug Information

#### 1. Missing Graph API Credentials
```
================================================================================
🖼️  RANGE IMAGE DEBUG: WARNING
📍 range_images - warning:45
💬 ⚠️ No Graph API credentials provided - range image extraction disabled
================================================================================
```
**Solution**: Ensure your Graph API credentials are properly configured.

#### 2. Sheet Not Found
```
================================================================================
🖼️  RANGE IMAGE DEBUG: ERROR
📍 range_images - error:47
💬 ❌ EXPORT FAILED: client_summary_table
   💥 Error: Sheet 'Order Form' not found. Available: ['Sheet1', 'Data']
================================================================================
```
**Solution**: Check the sheet names in your Excel file match the configuration.

#### 3. Invalid Range Format
```
================================================================================
🖼️  RANGE IMAGE DEBUG: ERROR
📍 range_images - error:47
💬 ❌ RANGE VALIDATION FAILED
   📊 Sheet: Order Form
   📍 Range: A1:Invalid
   💥 Error: Invalid range format
================================================================================
```
**Solution**: Ensure ranges follow Excel format (e.g., "A1:E15").

### Export Progress Tracking
Monitor multi-range exports with visual progress indicators:

```
================================================================================
🖼️  RANGE IMAGE DEBUG: INFO
📍 range_images - info:39
💬 ✅ EXPORT PROGRESS [2/4] 50%
   ██████████░░░░░░░░░░
   🏷️  Current: word_search_table
   🎯 Status: SUCCESS
================================================================================
```

### Final Results Summary
At the end of processing, see a complete summary:

```
================================================================================
🖼️  RANGE IMAGE DEBUG: INFO
📍 range_images - info:39
💬 🏁 EXTRACTION COMPLETED
   ✅ Successful: 3/4
   📁 Output Files: ['client_summary_table', 'word_search_table', 'quarterly_results_chart']
================================================================================
```

## Integration with Existing Workflows

### With Regular Debug Mode
The range image debug mode works alongside regular debug mode:
- Use `--debug` for general application debugging
- Use `--debug-range-images` specifically for range image issues
- Combine both flags for comprehensive debugging

### API Development
When developing against the API, the enhanced logging will appear in your server console when range image operations are performed via API endpoints.

### Production Use
In production, you can enable range image debugging temporarily by restarting the server with the `--debug-range-images` flag to troubleshoot specific range image issues without full debug verbosity.

## Advanced Usage

### Programmatic Debugging
You can also enable range image debug mode programmatically:

```python
from src.utils.range_image_logger import setup_range_image_debug_mode
import logging

# Enable enhanced range image debugging
setup_range_image_debug_mode(enabled=True, level=logging.DEBUG)

# Your range image processing code here...
```

### Custom Log Levels
Adjust the debug level for different verbosity:

```python
# More verbose debugging
setup_range_image_debug_mode(enabled=True, level=logging.DEBUG)

# Less verbose (INFO level only)
setup_range_image_debug_mode(enabled=True, level=logging.INFO)
```

## Summary

The enhanced range image debugging system provides:
- ✅ **Clear visibility** into range image processing steps
- ✅ **Noise reduction** from verbose PowerPoint processor logs  
- ✅ **Visual formatting** for easy identification of range image logs
- ✅ **Detailed progress tracking** for multi-range exports
- ✅ **Comprehensive error reporting** with actionable information
- ✅ **Easy activation** via command-line flags

No more searching through thousands of DEBUG lines to find your range image processing information!