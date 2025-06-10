# Excel to PowerPoint Merger - Product Requirements Document

## 1. Executive Summary

### 1.1 Product Overview
The Excel to PowerPoint Merger is a cloud-based API service that automatically extracts structured data from Excel spreadsheets and merges it into PowerPoint presentation templates using Jinja-style merge fields. The solution will be deployed as a Google Cloud Function and integrate with Zoho WorkDrive for file storage.

### 1.2 Business Objectives
- **Automation**: Eliminate manual data entry from Excel to PowerPoint presentations
- **Consistency**: Ensure standardized presentation formatting across the organization
- **Efficiency**: Reduce time spent on repetitive presentation creation tasks
- **Scalability**: Handle multiple concurrent requests via API
- **Integration**: Seamlessly work within the existing Zoho One ecosystem

### 1.3 Success Metrics
- **Performance**: Process typical files (1-5MB Excel, 1-10MB PPTX) within 30 seconds
- **Reliability**: 99.5% uptime and successful processing rate
- **Accuracy**: 100% data mapping accuracy for properly configured templates
- **Adoption**: Replace manual presentation creation workflows

## 2. Product Scope

### 2.1 In Scope
- Excel data extraction from multiple sheet types and table structures
- PowerPoint template processing with text and image merge fields
- Dynamic table detection and flexible data mapping
- Image extraction and placement from Excel files
- RESTful API with file upload/download capabilities
- Integration with Zoho WorkDrive for file storage
- Configuration-driven data mapping system
- Comprehensive error handling and logging

### 2.2 Out of Scope (Future Phases)
- Real-time collaboration features
- Advanced PowerPoint animations or transitions
- Multi-language template support
- Batch processing of multiple files
- User interface (web dashboard)
- Advanced Excel formula evaluation
- PowerPoint template creation tools

## 3. User Stories and Use Cases

### 3.1 Primary User Personas
- **Business Analysts**: Generate client reports with data from Excel analyses
- **Project Managers**: Create status presentations from project tracking spreadsheets
- **Sales Teams**: Produce client proposals with standardized templates
- **Automation Engineers**: Integrate presentation generation into existing workflows

### 3.2 Core User Stories

#### 3.2.1 Basic Data Merge
**As a** business analyst  
**I want to** upload an Excel file and PowerPoint template via API  
**So that** I can automatically generate a presentation with current data  

**Acceptance Criteria:**
- API accepts Excel (.xlsx) and PowerPoint (.pptx) files
- Successfully extracts data from predefined table structures
- Replaces all merge fields with corresponding data
- Returns completed presentation file
- Process completes within acceptable time limits

#### 3.2.2 Multi-Table Excel Processing
**As a** project manager  
**I want to** extract data from multiple tables within a single Excel sheet  
**So that** I can populate different sections of my presentation template  

**Acceptance Criteria:**
- Detects and extracts data from multiple subtables on one sheet
- Handles variable row counts in each table
- Maps each table to appropriate merge fields
- Maintains data relationships and structure

#### 3.2.3 Image Integration
**As a** sales representative  
**I want to** include company logos and product images from Excel into presentations  
**So that** my proposals are visually consistent and professional  

**Acceptance Criteria:**
- Extracts embedded images from Excel files
- Places images in designated placeholders in PowerPoint
- Maintains image quality and aspect ratios
- Supports common image formats (PNG, JPEG, GIF, WebP)

#### 3.2.4 Flexible Configuration
**As a** automation engineer  
**I want to** configure data mapping rules for different Excel layouts  
**So that** I can handle various spreadsheet formats without code changes  

**Acceptance Criteria:**
- Accepts JSON configuration for data extraction rules
- Dynamically locates table headers and data ranges
- Supports custom column name mappings
- Handles different table orientations and structures

## 4. Functional Requirements

### 4.1 Core API Functionality

#### 4.1.1 File Processing Endpoint
- **Endpoint**: `POST /api/v1/merge`
- **Input**: Excel file, PowerPoint template, optional configuration JSON
- **Output**: Completed PowerPoint file
- **Authentication**: API key-based authentication
- **Rate Limiting**: 100 requests per hour per API key

#### 4.1.2 Configuration Management
- **Endpoint**: `POST/GET /api/v1/config`
- **Functionality**: Store and retrieve data mapping configurations
- **Validation**: Validate configuration schema before processing
- **Versioning**: Support multiple configuration versions

#### 4.1.3 Status and Health Checks
- **Endpoint**: `GET /api/v1/health`
- **Functionality**: Return service health and processing statistics
- **Monitoring**: Integration with Cloud Monitoring

### 4.2 Data Extraction Engine

#### 4.2.1 Dynamic Table Detection
- Search for table headers using configurable criteria:
  - Text contains matching
  - First non-empty row detection
  - Regular expression patterns
- Support multiple search ranges per sheet
- Handle nested and adjacent tables

#### 4.2.2 Flexible Data Extraction
- **Table Types Supported**:
  - Standard vertical tables (headers at top)
  - Horizontal key-value pairs
  - Tables with embedded images
- **End Condition Detection**:
  - Consecutive empty rows
  - Specific text markers
  - Maximum row limits
- **Data Normalization**:
  - Column name standardization (spaces → underscores)
  - Whitespace trimming
  - Empty cell handling

#### 4.2.3 Image Processing
- Extract embedded images from Excel files
- Determine image positioning relative to table data
- Support multiple image formats
- Temporary image storage for processing

#### 4.2.4 Temporary File Management
- **Configurable Cleanup**: Environment-based control for temporary file cleanup
- **Development Mode**: Option to retain temporary files for debugging and inspection
- **Production Mode**: Automatic cleanup to prevent storage bloat
- **Error Handling**: Configurable retention of temporary files when processing errors occur
- **Delayed Cleanup**: Configurable retention period before automatic cleanup
- **Runtime Override**: API parameters to control cleanup behavior per request

### 4.3 PowerPoint Template Processing

#### 4.3.1 Merge Field Replacement
- Support Jinja-style merge fields: `{{field_name}}`
- Handle text content replacement
- Maintain original formatting when possible
- Support nested object references: `{{table.0.field_name}}`

#### 4.3.2 Image Placeholder Handling
- Replace text placeholders with actual images
- Maintain placeholder dimensions and positioning
- Automatic image resizing to fit containers
- Support for multiple images per template

#### 4.3.3 Template Validation
- Verify merge field syntax before processing
- Check for missing required fields
- Provide detailed error messages for invalid templates

## 5. Technical Requirements

### 5.1 Architecture Overview
- **Deployment**: Google Cloud Function (Gen 2)
- **Runtime**: Python 3.11
- **Storage**: Zoho WorkDrive integration
- **Monitoring**: Google Cloud Monitoring and Logging

### 5.2 Performance Requirements
- **Processing Time**: 95% of requests complete within 30 seconds
- **File Size Limits**: 
  - Excel files: up to 50MB
  - PowerPoint templates: up to 100MB
  - Generated presentations: up to 150MB
- **Concurrent Requests**: Support up to 10 simultaneous requests
- **Memory Usage**: Optimize for Cloud Function memory limits (8GB max)

### 5.3 Data Format Requirements

#### 5.3.1 Excel File Support
- **Formats**: .xlsx (OpenXML format)
- **Sheets**: Multiple sheets per file
- **Tables**: Multiple tables per sheet
- **Images**: Embedded images in supported formats
- **Data Types**: Text, numbers, dates, formulas (values only)

#### 5.3.2 PowerPoint Template Support
- **Formats**: .pptx (OpenXML format)
- **Slides**: Multiple slides per template
- **Merge Fields**: Text-based placeholders
- **Images**: Text placeholder replacement
- **Formatting**: Preserve existing slide formatting

#### 5.3.3 Configuration Schema
```json
{
  "version": "1.0",
  "sheet_configs": {
    "sheet_name": {
      "subtables": [
        {
          "name": "table_identifier",
          "type": "table|key_value_pairs|table_with_images",
          "header_search": {
            "method": "contains_text|first_non_empty_row|regex",
            "text": "search_string",
            "column": "A",
            "search_range": "A1:A20"
          },
          "data_extraction": {
            "orientation": "vertical|horizontal",
            "headers_row_offset": 0,
            "data_start_row_offset": 1,
            "max_rows": 100,
            "end_condition": {
              "method": "empty_rows|text_marker|max_rows",
              "consecutive_empty": 2,
              "marker_text": "END"
            },
            "column_mappings": {
              "Excel Header": "json_key"
            }
          }
        }
      ]
    }
  },
  "global_settings": {
    "normalize_keys": true,
    "image_output_dir": "extracted_images",
    "empty_cell_value": "",
    "temp_file_cleanup": {
      "enabled": true,
      "delay_seconds": 300,
      "keep_on_error": true,
      "development_mode": false
    }
  }
}
```

### 5.4 Integration Requirements

#### 5.4.1 Zoho WorkDrive Integration
- **Authentication**: OAuth 2.0 with service account
- **File Operations**: Upload, download, delete temporary files
- **Folder Structure**: Organized temporary storage for processing
- **Error Handling**: Graceful handling of API limits and errors

#### 5.4.2 Security Requirements
- **API Authentication**: API key validation
- **File Security**: Temporary file cleanup after processing
- **Data Privacy**: No persistent storage of user data
- **Access Control**: Proper IAM roles for Cloud Function

#### 5.4.3 Temporary File Management
- **Environment Variables**: Support for deployment-specific cleanup configuration
- **Development Environment**: 
  - `DEVELOPMENT_MODE=true`: Disable automatic cleanup for debugging
  - `CLEANUP_TEMP_FILES=false`: Override cleanup behavior
  - `TEMP_FILE_RETENTION_SECONDS=3600`: Extended retention for analysis
- **Production Environment**:
  - `DEVELOPMENT_MODE=false`: Enable production cleanup policies
  - `CLEANUP_TEMP_FILES=true`: Automatic cleanup enabled
  - `TEMP_FILE_RETENTION_SECONDS=60`: Quick cleanup for resource management
- **Runtime Parameters**:
  - `keep_temp_files` (API parameter): Per-request cleanup override
  - `temp_retention_seconds` (API parameter): Custom retention period
- **Error Handling**: Configurable file retention when processing fails for debugging

## 6. Non-Functional Requirements

### 6.1 Reliability
- **Uptime**: 99.5% availability
- **Error Handling**: Comprehensive error catching and logging
- **Recovery**: Automatic retry for transient failures
- **Monitoring**: Real-time health checks and alerting

### 6.2 Scalability
- **Auto-scaling**: Automatic Cloud Function scaling based on demand
- **Resource Management**: Efficient memory and CPU usage
- **Concurrent Processing**: Handle multiple simultaneous requests
- **Load Testing**: Support peak loads up to 50 requests/minute

### 6.3 Maintainability
- **Code Quality**: Comprehensive documentation and testing
- **Logging**: Detailed request/response logging
- **Monitoring**: Performance metrics and error tracking
- **Deployment**: Automated CI/CD pipeline

### 6.4 Security
- **Data Encryption**: In-transit and at-rest encryption
- **Access Control**: API key-based authentication
- **File Validation**: Input file format and size validation
- **Vulnerability Management**: Regular security updates

## 7. API Specification

### 7.1 Merge Endpoint
```
POST /api/v1/merge
Content-Type: multipart/form-data

Parameters:
- excel_file (file): Excel spreadsheet (.xlsx)
- pptx_template (file): PowerPoint template (.pptx)
- config (optional, JSON): Data extraction configuration
- output_filename (optional, string): Custom output filename
- keep_temp_files (optional, boolean): Override temporary file cleanup (default: false)
- temp_retention_seconds (optional, integer): Custom retention period for temp files (default: 300)

Response:
200 OK
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation
Content-Disposition: attachment; filename="merged_presentation.pptx"

Error Responses:
400 Bad Request: Invalid file format or configuration
401 Unauthorized: Invalid API key
413 Payload Too Large: File size exceeds limits
422 Unprocessable Entity: Data extraction or merge errors
500 Internal Server Error: Processing failure
```

### 7.2 Configuration Endpoint
```
POST /api/v1/config
Content-Type: application/json

{
  "name": "config_name",
  "description": "Configuration description",
  "config": { /* configuration schema */ }
}

GET /api/v1/config/{config_name}
Response: Configuration JSON

PUT /api/v1/config/{config_name}
Update existing configuration

DELETE /api/v1/config/{config_name}
Remove configuration
```

### 7.3 Health Check Endpoint
```
GET /api/v1/health

Response:
{
  "status": "healthy",
  "version": "1.0.0",
  "uptime": 3600,
  "processed_requests": 1234,
  "error_rate": 0.02
}
```

## 8. Error Handling and Edge Cases

### 8.1 File Processing Errors
- **Invalid Excel format**: Return 400 with specific error message
- **Corrupted PowerPoint template**: Return 422 with validation details
- **Missing merge fields**: Log warning, continue with available data
- **Image extraction failure**: Log error, continue without images

### 8.2 Data Extraction Errors
- **Table not found**: Return 422 with configuration suggestions
- **Empty data sets**: Return 200 with warning in response headers
- **Column mapping mismatch**: Use auto-generated mappings as fallback
- **Invalid data types**: Convert to string representation

### 8.4 Temporary File Management Errors
- **Cleanup failures**: Log errors but don't fail main process
- **Storage quota exceeded**: Implement emergency cleanup procedures
- **Permission errors**: Graceful handling of file system permission issues
- **Development mode**: Preserve temp files on all errors for debugging

## 9. Testing Strategy

### 9.1 Unit Testing
- Data extraction functions
- PowerPoint processing functions
- Configuration validation
- Error handling scenarios

### 9.2 Integration Testing
- End-to-end API workflows
- Zoho WorkDrive integration
- File upload/download scenarios
- Error condition testing

### 9.4 Temporary File Management Testing
- **Cleanup behavior**: Verify files are cleaned up in production mode
- **Development mode**: Confirm files are retained for debugging
- **Error scenarios**: Test file retention when processing fails
- **Resource management**: Monitor storage usage during processing
- **Configuration testing**: Verify environment variable and runtime parameter behavior

## 10. Deployment and Operations

### 10.1 Deployment Pipeline
1. **Development**: Local testing with sample files
2. **Staging**: Cloud Function deployment with test data
3. **Production**: Gradual rollout with monitoring
4. **Rollback**: Automated rollback procedures

### 10.2 Monitoring and Alerting
- **Request/Response Metrics**: Success rate, latency, error types
- **Resource Usage**: Memory, CPU, storage utilization
- **Integration Health**: Zoho API response times and errors
- **Alert Thresholds**: Error rate >5%, latency >45 seconds

### 10.3 Maintenance
- **Log Retention**: 30 days for detailed logs, 90 days for metrics
- **Security Updates**: Monthly dependency updates
- **Performance Optimization**: Quarterly performance reviews
- **Capacity Planning**: Monthly usage trend analysis

## 11. Future Enhancements (Out of Scope)

### 11.1 Phase 2 Features
- **Batch Processing**: Handle multiple file pairs simultaneously
- **Template Library**: Store and manage reusable templates
- **Advanced Formatting**: Support for complex PowerPoint features
- **Real-time Preview**: Generate preview before final processing

### 11.2 Phase 3 Features
- **Web Dashboard**: User interface for configuration management
- **Workflow Integration**: Zapier/Microsoft Power Automate connectors
- **Advanced Analytics**: Usage reporting and optimization suggestions
- **Multi-tenant Architecture**: Support for multiple organizations

---

**Document Version**: 1.0  
**Last Updated**: June 10, 2025  
**Review Date**: July 10, 2025  
**Stakeholders**: Development Team, Product Management, Operations