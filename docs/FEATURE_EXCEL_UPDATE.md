# Feature: Excel File Update

## Overview
Add capability to update Excel files with new data through a REST API endpoint. Users can modify cell values, insert images, and update table data using flexible configuration mapping that supports both direct cell addressing and search-based location methods.

## Problem Statement
Currently, the system only extracts data from Excel files but cannot update them. Users need to:
- Update specific cells with new values (text, numbers, images)
- Modify table data with batch updates
- Handle both direct cell addressing and search-based cell location
- Receive diagnostic information about update operations
- Get comprehensive error handling with fallback values

## Proposed Solution
Implement Excel file update functionality through a new REST API endpoint that accepts Excel files, update data, and configuration mappings, returning updated Excel files with diagnostic information.

## Technical Implementation Plan

### Phase 1: Core Update Engine (MVP)

#### 1.1 New API Endpoint
```http
POST /api/v1/update
Content-Type: multipart/form-data

Parameters:
- excel_file: Excel file to update
- update_data: JSON data for updates
- config: JSON configuration mapping
```

#### 1.2 Core Update Module
Create new module `src/excel_updater.py`:

**Key Classes:**
- `ExcelUpdater`: Main class for Excel file updates
- `UpdateLocation`: Helper class for location resolution
- `UpdateLogger`: Diagnostic logging and error tracking

**Core Methods:**
```python
class ExcelUpdater:
    def update_excel(self, update_data: Dict, config: Dict) -> str
    def _find_update_location(self, sheet, header_search_config) -> Dict
    def _update_subtable(self, sheet, subtable_config, update_data) -> None
    def _update_cell(self, sheet, cell_address, value, field_type) -> bool
    def _insert_image(self, sheet, cell_address, image_data) -> bool
    def _add_update_log_sheet(self) -> None
```

#### 1.3 Configuration Schema Extension
Support both existing extraction methods and new update-specific options:

```json
{
  "sheet_configs": {
    "Sheet1": {
      "subtables": [
        {
          "name": "client_info",
          "type": "key_value_pairs",
          "header_search": {
            "method": "cell_address",
            "cell": "B14"
          },
          "data_update": {
            "orientation": "horizontal",
            "headers_row_offset": 0,
            "data_row_offset": 1,
            "column_mappings": {
              "B14": {"name": "client_name", "type": "text"},
              "C14": {"name": "logo", "type": "image"}
            }
          }
        }
      ]
    }
  }
}
```

#### 1.4 Mixed Header Search Methods

**Method 1: Direct Cell Address**
```json
{
  "header_search": {
    "method": "cell_address",
    "cell": "B14"
  },
  "data_update": {
    "column_mappings": {
      "B14": {"name": "field_name", "type": "text"}
    }
  }
}
```

**Method 2: Text Search with Offsets**
```json
{
  "header_search": {
    "method": "contains_text",
    "text": "Client Information",
    "column": "A",
    "search_range": "A1:A20"
  },
  "data_update": {
    "headers_row_offset": 1,
    "data_row_offset": 2,
    "column_mappings": {
      "Client Name": {"name": "client_name", "type": "text"}
    }
  }
}
```

#### 1.5 Data Type Support

**Text Fields:**
```python
cell.value = str(value)
```

**Number Fields:**
```python
cell.value = float(value) if value else 0
```

**Image Fields:**
```python
# Support multiple input formats:
# 1. Base64: "data:image/png;base64,iVBOR..."
# 2. URL: "https://example.com/image.jpg"
# 3. Text fallback: Plain text if image processing fails

# Insert image above target cell
img = PILImage.open(io.BytesIO(image_bytes))
excel_img = ExcelImage(img)
excel_img.anchor = f"{get_column_letter(col)}{row-1}"
sheet.add_image(excel_img)
```

#### 1.6 Table Data Updates
For table-type subtables with vertical orientation:

```json
{
  "header_search": {
    "method": "cell_address", 
    "cell": "A20"
  },
  "data_update": {
    "orientation": "vertical",
    "data_row_offset": 1,
    "column_mappings": {
      "A": {"name": "item_name", "type": "text"},
      "B": {"name": "price", "type": "number"}
    }
  }
}
```

Input data format:
```json
{
  "items_table": [
    {"item_name": "Product A", "price": 100},
    {"item_name": "Product B", "price": 200}
  ]
}
```

### Phase 2: Advanced Features

#### 2.1 Enhanced Error Handling
- **Partial Success Processing**: Continue updates even if some cells fail
- **Error Cell Marking**: Insert `!#VALUE` for failed operations
- **Comprehensive Logging**: Track all operations in `update_log` sheet

#### 2.2 Image Processing Pipeline
```python
def _process_image_data(self, image_data: str) -> bytes:
    """Process image with multiple input format support"""
    
    if image_data.startswith('data:image'):
        # Base64 encoded image
        return self._decode_base64_image(image_data)
    elif image_data.startswith('http'):
        # Download from URL
        return self._download_image_with_timeout(image_data)
    else:
        # Text fallback - return None to use text value
        return None
```

#### 2.3 Image Compression Integration (Future)
**Cloudinary API Integration:**
```python
# Future implementation - placeholder for now
def _compress_image_if_needed(self, image_bytes: bytes, config: Dict) -> bytes:
    """
    Compress large images using Cloudinary API
    
    Features:
    - Automatic compression for images above configurable size threshold
    - Configurable quality settings
    - Automatic cleanup to prevent storage overflow
    - Can be disabled via config
    """
    
    compression_config = config.get("image_compression", {})
    enabled = compression_config.get("enabled", True)
    size_threshold_mb = compression_config.get("threshold_mb", 2)
    
    if not enabled:
        return image_bytes
        
    # Check image size
    image_size_mb = len(image_bytes) / (1024 * 1024)
    
    if image_size_mb > size_threshold_mb:
        # TODO: Implement Cloudinary compression
        # - Upload to Cloudinary with compression settings
        # - Download compressed version
        # - Schedule cleanup job
        # - Return compressed bytes
        pass
        
    return image_bytes
```

**Configuration Options:**
```json
{
  "global_settings": {
    "image_compression": {
      "enabled": true,
      "threshold_mb": 2,
      "quality": 80,
      "format": "auto",
      "cleanup_delay_hours": 24
    }
  }
}
```

#### 2.4 Update Log Sheet Structure
Automatic creation of diagnostic sheet:

| Column | Field | Description |
|--------|-------|-------------|
| A | Timestamp | ISO timestamp of operation |
| B | Operation | Type of update (cell_update, image_insert, table_update) |
| C | Cell/Range | Target cell or range address |
| D | Status | SUCCESS, ERROR, WARNING |
| E | Details | Error messages or success details |
| F | Original Value | Value before update (for rollback reference) |
| G | New Value | Value after update |

### Phase 3: Production Features

#### 3.1 Advanced Validation
```python
def _validate_update_request(self, update_data: Dict, config: Dict) -> List[str]:
    """Comprehensive validation with detailed error reporting"""
    
    errors = []
    
    # Validate configuration structure
    errors.extend(self._validate_config_structure(config))
    
    # Validate data types match configuration
    errors.extend(self._validate_data_types(update_data, config))
    
    # Validate cell addresses exist in sheets
    errors.extend(self._validate_cell_addresses(config))
    
    # Validate image data formats
    errors.extend(self._validate_image_data(update_data))
    
    return errors
```

#### 3.2 Performance Optimization
- **Batch Processing**: Group cell updates to minimize Excel file saves
- **Memory Management**: Stream large images, cleanup temporary data
- **Concurrent Safety**: Handle multiple simultaneous updates safely
- **Image Caching**: Cache processed images to avoid reprocessing

#### 3.3 Rollback Support
```python
def _create_backup_sheet(self) -> str:
    """Create backup of original data before updates"""
    
    backup_sheet_name = f"backup_{int(time.time())}"
    # Store original values for potential rollback
```

## Implementation Dependencies

### Required Libraries
```python
# Core (already in use)
openpyxl>=3.1.0          # Excel file manipulation
Pillow>=10.0.0           # Image processing
Flask>=2.3.0             # Web framework

# New dependencies
requests>=2.31.0         # Image URL downloads (already in use)
validators>=0.20.0       # URL and data validation

# Future (Image compression)
cloudinary>=1.34.0       # Image compression service
```

### System Requirements
- Sufficient disk space for temporary image files
- Memory for processing large images and Excel files
- Network access for image URL downloads
- Optional: Cloudinary account for compression features

## Configuration Migration

### Backward Compatibility
- Existing extraction configurations continue to work unchanged
- Update feature is additive - no breaking changes to current API
- New `data_update` section is optional in configurations

### Migration Path
1. Add `data_update` sections to existing configurations
2. Test update operations on copies of production files
3. Implement gradual rollout with comprehensive logging
4. Monitor performance and error rates

## API Response Format

### Success Response
```json
{
  "success": true,
  "message": "Excel file updated successfully",
  "summary": {
    "total_operations": 15,
    "successful_operations": 14,
    "failed_operations": 1,
    "warnings": 2
  },
  "details": {
    "updated_sheets": ["Sheet1", "Sheet2"],
    "log_sheet_created": true,
    "processing_time_ms": 1250
  }
}
```

### Error Response
```json
{
  "success": false,
  "error": "Validation failed",
  "details": [
    "Cell address 'ZZ999' does not exist in sheet 'Sheet1'",
    "Invalid image data format for field 'logo'"
  ],
  "partial_results": {
    "completed_operations": 5,
    "failed_at_operation": "image_insert_B14"
  }
}
```

## Testing Strategy

### Unit Tests
```python
class TestExcelUpdater:
    def test_cell_address_method(self):
        """Test direct cell addressing"""
        
    def test_contains_text_method(self):
        """Test text search with offsets"""
        
    def test_mixed_configuration(self):
        """Test both methods in same config"""
        
    def test_image_processing(self):
        """Test base64 and URL image handling"""
        
    def test_error_handling(self):
        """Test graceful error recovery"""
        
    def test_update_log_creation(self):
        """Test diagnostic logging"""
```

### Integration Tests
- End-to-end file update workflows
- Large file processing performance
- Concurrent update handling
- Image download timeout scenarios
- Invalid configuration handling

### Test Data
- Sample Excel files with various layouts
- Test images in multiple formats
- Edge cases: merged cells, protected sheets, formula cells
- Invalid data scenarios for error handling

## Documentation Requirements

### User Documentation
- API endpoint specification and examples
- Configuration guide for update mappings
- Image format requirements and limitations
- Error handling and troubleshooting guide
- Best practices for large file processing

### Developer Documentation
- Architecture overview of update pipeline
- Extension guide for new data types
- Image compression integration guide
- Performance tuning recommendations

## Success Metrics

### Functional Goals
- [ ] Successfully update Excel cells with text and numeric data
- [ ] Insert images from base64 and URL sources
- [ ] Support both cell_address and contains_text location methods
- [ ] Handle table data updates with batch processing
- [ ] Generate comprehensive diagnostic logs

### Performance Goals
- [ ] Process files up to 50MB in <10 seconds
- [ ] Handle concurrent updates (5+ simultaneous requests)
- [ ] Memory usage scales linearly with file size
- [ ] Image downloads complete within 30 seconds timeout

### Quality Goals
- [ ] 95%+ successful update rate for valid requests
- [ ] Comprehensive error logging and recovery
- [ ] Zero data corruption in target Excel files
- [ ] Complete audit trail for all operations

## Future Enhancements

### Advanced Features
- Formula preservation and recalculation
- Conditional formatting updates
- Chart data source updates
- Multi-sheet transaction support (all-or-nothing updates)
- Template-based bulk updates

### Image Processing
- Advanced image compression with Cloudinary
- Image format conversion (PNG to JPEG, etc.)
- Automatic image resizing for cell dimensions
- Image optimization for web and print

### Performance
- Streaming updates for very large files
- Background processing for time-intensive operations
- Caching layer for frequently updated files
- Delta updates (only changed cells)

## Risk Assessment

### Technical Risks
- **Large image processing**: Memory consumption for high-resolution images
- **Excel file corruption**: Malformed updates could corrupt files
- **Concurrent access**: Multiple updates to same file simultaneously
- **Image download failures**: Network timeouts and invalid URLs

### Mitigation Strategies
- Implement file size limits and image compression
- Create backups before updates and validate file integrity
- Use file locking and queue-based processing
- Robust timeout handling and fallback to text values
- Comprehensive testing with edge cases

## Implementation Timeline

### Week 1-2: Core Foundation
- Implement ExcelUpdater class with basic cell updates
- Add support for both header search methods
- Create update configuration validation
- Basic error handling and logging

### Week 3-4: Image Processing
- Implement image insertion from base64 and URLs
- Add image format validation and error handling
- Create update log sheet functionality
- Enhanced error recovery mechanisms

### Week 5-6: API Integration
- Create /api/v1/update endpoint
- Integrate with existing Flask application
- Add request validation and response formatting
- Implement file handling and cleanup

### Week 7-8: Testing and Polish
- Comprehensive unit and integration testing
- Performance optimization and memory management
- Documentation and examples
- Production deployment preparation

## Acceptance Criteria

### Must Have
- [ ] Update Excel cells using both cell_address and contains_text methods
- [ ] Support text, number, and image data types
- [ ] Handle table data with batch updates
- [ ] Generate diagnostic update_log sheet
- [ ] Graceful error handling with partial success
- [ ] Return updated Excel file in API response

### Should Have
- [ ] Image download from URLs with timeout handling
- [ ] Configuration validation with detailed error messages
- [ ] Performance optimization for large files
- [ ] Comprehensive logging and monitoring

### Could Have
- [ ] Image compression integration (Cloudinary placeholder)
- [ ] Advanced image processing features
- [ ] Rollback functionality
- [ ] Multi-sheet transaction support