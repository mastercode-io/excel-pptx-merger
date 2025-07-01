# Feature: Excel Range to Image Export

## Overview
Add capability to export Excel cell ranges as images and insert them into PowerPoint presentations, preserving formatting, styling, and visual appearance exactly as they appear in Excel.

## Problem Statement
Currently, the system extracts individual cell values and images from Excel, but cannot capture formatted table ranges as visual blocks. Users often need to include Excel tables/ranges in presentations while maintaining:
- Cell formatting (borders, colors, fonts)
- Column widths and row heights  
- Visual styling and layout
- Complex table structures

## Proposed Solution
Implement Excel range-to-image export functionality that allows users to specify cell ranges in configuration and have them automatically converted to images for PowerPoint insertion.

## Technical Implementation Plan

### Phase 1: Core Range Export (MVP)

#### 1.1 Configuration Schema Extension
Extend the existing configuration to support range image definitions:

```json
{
  "range_images": [
    {
      "field_name": "summary_table",
      "sheet_name": "Data", 
      "range": "A1:E15",
      "include_headers": true,
      "output_format": "png",
      "dpi": 150,
      "fit_to_content": true
    },
    {
      "field_name": "quarterly_results",
      "sheet_name": "Q4_Data",
      "range": "B2:H20", 
      "include_headers": false,
      "output_format": "png",
      "dpi": 300
    }
  ]
}
```

#### 1.2 Excel Range Export Module
Create new module `src/excel_range_exporter.py`:

**Key Classes:**
- `ExcelRangeExporter`: Main class for range-to-image conversion
- `RangeImageConfig`: Configuration model for range definitions
- `RangeImageResult`: Result object containing image path and metadata

**Core Methods:**
```python
class ExcelRangeExporter:
    def export_range_as_image(self, workbook_path: str, config: RangeImageConfig) -> RangeImageResult
    def _validate_range(self, sheet, range_str: str) -> bool
    def _calculate_range_dimensions(self, sheet, range_str: str) -> Tuple[int, int]
    def _export_with_xlwings(self, workbook_path: str, config: RangeImageConfig) -> str
    def _export_with_openpyxl_pil(self, workbook_path: str, config: RangeImageConfig) -> str
```

#### 1.3 Library Integration Options

**Option A: xlwings (Recommended for Windows/Mac with Excel installed)**
```python
import xlwings as xw

def export_range_xlwings(workbook_path, sheet_name, range_str, output_path):
    app = xw.App(visible=False)
    wb = app.books.open(workbook_path)
    sheet = wb.sheets[sheet_name]
    range_obj = sheet.range(range_str)
    range_obj.api.CopyPicture(Format=2)  # xlBitmap format
    # Export to file
    wb.close()
    app.quit()
```

**Option B: openpyxl + PIL (Cross-platform fallback)**
```python
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

def export_range_openpyxl(workbook_path, sheet_name, range_str, output_path):
    wb = load_workbook(workbook_path)
    sheet = wb[sheet_name]
    # Render cells manually using PIL
    # More complex but platform-independent
```

#### 1.4 Integration with Excel Processor
Extend `src/excel_processor.py`:

```python
class ExcelProcessor:
    def __init__(self, file_path: str):
        self.range_exporter = ExcelRangeExporter()
    
    def extract_data_and_ranges(self, config: Dict[str, Any]) -> Dict[str, Any]:
        # Existing data extraction
        data = self._extract_data(config)
        
        # New range image extraction
        if "range_images" in config:
            range_images = self._extract_range_images(config["range_images"])
            data["_range_images"] = range_images
        
        return data
    
    def _extract_range_images(self, range_configs: List[Dict]) -> Dict[str, str]:
        range_images = {}
        for range_config in range_configs:
            image_path = self.range_exporter.export_range_as_image(
                self.file_path, 
                RangeImageConfig(**range_config)
            )
            range_images[range_config["field_name"]] = image_path
        return range_images
```

#### 1.5 PowerPoint Integration
Extend `src/pptx_processor.py` to handle range images:

```python
class PowerPointProcessor:
    def _is_range_image_field(self, field_name: str, data: Dict[str, Any]) -> bool:
        return "_range_images" in data and field_name in data["_range_images"]
    
    def _get_range_image_path(self, field_name: str, data: Dict[str, Any]) -> Optional[str]:
        if self._is_range_image_field(field_name, data):
            return data["_range_images"][field_name]
        return None
```

### Phase 2: Advanced Features

#### 2.1 Dynamic Range Detection
- Auto-detect table boundaries based on data
- Smart range expansion/contraction
- Header row detection

#### 2.2 Styling Options
```json
{
  "range_images": [
    {
      "field_name": "styled_table",
      "sheet_name": "Data",
      "range": "A1:E15",
      "styling": {
        "zoom_factor": 1.2,
        "border_style": "medium",
        "background_color": "#FFFFFF",
        "remove_gridlines": true,
        "custom_width": 800,
        "custom_height": 600
      }
    }
  ]
}
```

#### 2.3 Multiple Format Support
- PNG (default)
- JPEG (smaller files)
- SVG (scalable, if supported)
- PDF (high quality)

#### 2.4 Performance Optimization
- Image caching based on range content hash
- Parallel processing for multiple ranges
- Memory optimization for large ranges

### Phase 3: Enhanced Integration

#### 3.1 Template-Based Range Selection
Allow PowerPoint templates to define range requirements:
```
{{range_image:summary_table:A1:E15}}
{{range_image:quarterly_data:auto_detect}}
```

#### 3.2 Range Validation and Feedback
- Validate ranges exist and contain data
- Provide warnings for empty ranges
- Suggest optimal DPI based on range size

#### 3.3 Error Handling and Fallbacks
- Graceful degradation if range export fails
- Fallback to table recreation using cell data
- Detailed error reporting

## Implementation Dependencies

### Required Libraries
```python
# Primary (choose one)
xlwings>=0.24.0          # For Windows/Mac with Excel
# OR
openpyxl>=3.0.0          # Cross-platform
Pillow>=8.0.0            # For image generation fallback

# Existing
python-pptx>=0.6.21      # Already in use
pandas>=1.3.0            # Already in use
```

### System Requirements
- **xlwings option**: Excel installation required (Windows/Mac)
- **openpyxl option**: Cross-platform, no Excel required
- Sufficient disk space for temporary image files
- Memory for processing large ranges

## Configuration Migration

### Backward Compatibility
- Existing configurations continue to work unchanged
- Range images are additive feature
- No breaking changes to current API

### Migration Path
1. Add `range_images` section to existing configs
2. Update field types to include "range_image" type
3. Update PowerPoint templates with range image placeholders

## Testing Strategy

### Unit Tests
- Range validation logic
- Image generation with mock data
- Configuration parsing and validation
- Error handling for invalid ranges

### Integration Tests
- End-to-end range export with real Excel files
- PowerPoint insertion with various image sizes
- Performance testing with large ranges
- Cross-platform compatibility testing

### Test Data
- Sample Excel files with various table formats
- PowerPoint templates with range image placeholders
- Edge cases: empty ranges, merged cells, complex formatting

## Documentation Requirements

### User Documentation
- Configuration guide for range images
- PowerPoint template setup instructions
- Troubleshooting guide for common issues
- Performance recommendations

### Developer Documentation
- API reference for new classes and methods
- Architecture overview of range export pipeline
- Extension guide for custom export formats

## Success Metrics

### Functional Goals
- [ ] Successfully export Excel ranges as images
- [ ] Preserve visual formatting and styling
- [ ] Integrate seamlessly with existing merge workflow
- [ ] Support both xlwings and openpyxl backends

### Performance Goals
- [ ] Export ranges up to 100x100 cells in <5 seconds
- [ ] Memory usage scales linearly with range size
- [ ] Support concurrent range exports
- [ ] Image file sizes optimized for quality vs. size

### Quality Goals
- [ ] 95%+ visual fidelity to original Excel appearance
- [ ] Robust error handling and recovery
- [ ] Cross-platform compatibility
- [ ] Comprehensive test coverage (>90%)

## Future Enhancements

### Advanced Features
- Interactive range selection UI
- Real-time preview of range exports
- Batch processing of multiple workbooks
- Integration with Excel online/cloud versions

### Export Options
- Vector format export (SVG/PDF)
- Custom themes and styling
- Watermarking and annotations
- Multi-page range support for large tables

## Risk Assessment

### Technical Risks
- **xlwings dependency**: Requires Excel installation, platform limitations
- **Image quality**: Balancing file size vs. visual fidelity  
- **Performance**: Large ranges may consume significant memory/time
- **Cross-platform**: Ensuring consistent behavior across OS

### Mitigation Strategies
- Provide both xlwings and openpyxl implementations
- Implement image compression and optimization
- Add memory monitoring and limits
- Extensive cross-platform testing

## Implementation Timeline

### Week 1-2: Foundation
- Design configuration schema
- Implement basic ExcelRangeExporter class
- Create unit test framework

### Week 3-4: Core Export
- Implement xlwings backend
- Implement openpyxl/PIL fallback backend
- Integration with ExcelProcessor

### Week 5-6: PowerPoint Integration
- Extend PowerPointProcessor for range images
- Update field detection and replacement logic
- End-to-end testing

### Week 7-8: Polish and Documentation
- Performance optimization
- Error handling improvements
- Documentation and examples
- Release preparation

## Acceptance Criteria

### Must Have
- [ ] Export Excel ranges to PNG images with preserved formatting
- [ ] Insert range images into PowerPoint via merge fields
- [ ] Support both xlwings and openpyxl backends
- [ ] Maintain backward compatibility with existing features
- [ ] Comprehensive error handling and logging

### Should Have  
- [ ] Configurable image quality and DPI settings
- [ ] Auto-detection of table boundaries
- [ ] Performance optimization for large ranges
- [ ] Cross-platform compatibility testing

### Could Have
- [ ] Multiple image format support (JPEG, SVG)
- [ ] Custom styling and theming options
- [ ] Batch processing capabilities
- [ ] Interactive range selection tools