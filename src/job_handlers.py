"""
Job Handlers - Endpoint to function mapping for job queue system

This module maps API endpoints to their corresponding handler functions,
allowing the job queue to call endpoint logic directly without HTTP requests.
"""

import logging
import json
from typing import Dict, Any, Callable, Optional, List
from .utils.exceptions import APIError

logger = logging.getLogger(__name__)


class MockRequest:
    """Mock Flask request object for job queue integration."""
    
    def __init__(self, job_payload: Dict[str, Any]):
        """Initialize mock request with job payload data."""
        # Determine if this should be treated as a JSON request
        self.is_json = True
        self._json_data = job_payload
        self.data = json.dumps(job_payload).encode('utf-8')
        
        # Initialize files dict - will be populated if base64 files are present
        self.files = {}
        
        # Initialize form dict for multipart simulation (not used in JSON mode)
        self.form = {}
        
        # Add headers dict for compatibility
        self.headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'JobQueue/1.0'
        }
        
        # Handle base64-encoded files in the payload
        if 'excel_file_base64' in job_payload:
            # For endpoints that expect 'excel_file' in files
            import base64
            import io
            excel_data = base64.b64decode(job_payload['excel_file_base64'])
            excel_file = io.BytesIO(excel_data)
            excel_file.filename = job_payload.get('excel_filename', 'excel_file.xlsx')
            # Note: We don't add to self.files for JSON requests
            # The PayloadParser will look for base64 data in JSON
        
        if 'pptx_file_base64' in job_payload:
            # For endpoints that expect 'pptx_file' in files
            import base64
            import io
            pptx_data = base64.b64decode(job_payload['pptx_file_base64'])
            pptx_file = io.BytesIO(pptx_data)
            pptx_file.filename = job_payload.get('pptx_filename', 'template.pptx')
            # Note: We don't add to self.files for JSON requests
    
    def get_json(self):
        """Return the JSON payload data."""
        return self._json_data
    
    @property
    def method(self):
        """Return the HTTP method (always POST for job requests)."""
        return 'POST'
    
    @property
    def path(self):
        """Return the request path."""
        return '/api/v1/job'


def extract_handler(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler for /api/v1/extract endpoint - Calls original endpoint with internal_data
    
    Args:
        payload: Request payload containing extraction parameters
        
    Returns:
        Extraction result dictionary
    """
    try:
        logger.info("🔄 Starting extract job processing - calling original endpoint with internal_data")
        
        # Import the original endpoint handler
        from . import main
        
        # Call the original endpoint handler with internal_data parameter
        result = main.extract_data_endpoint(internal_data=payload)
        
        # Handle tuple response (data, status_code)
        if isinstance(result, tuple):
            response_data, status_code = result
            if status_code == 200:
                return response_data
            else:
                raise APIError(f"Extract endpoint returned status {status_code}: {response_data}")
        
        return result
        
    except Exception as e:
        logger.exception(f"❌ Extract job failed: {str(e)}")
        raise APIError(f"Extract job failed: {str(e)}")


def merge_handler(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler for /api/v1/merge endpoint - Reuses original endpoint logic
    
    Args:
        payload: Request payload containing merge parameters
        
    Returns:
        Merged file result dictionary
    """
    try:
        logger.info("🔄 Starting merge job processing - calling original endpoint")
        
        # Import the original endpoint handler
        from . import main
        
        # Create mock request object
        mock_request = MockRequest(payload)
        
        # Create a test request context and override the request
        with main.app.test_request_context():
            # Temporarily replace the Flask request object
            import flask
            original_request = flask.request
            flask.request = mock_request
            
            try:
                # Call the original endpoint handler
                result = main.merge_files()
                
                # Handle tuple response (data, status_code)
                if isinstance(result, tuple):
                    response_data, status_code = result
                    if status_code == 200:
                        return response_data
                    else:
                        raise APIError(f"Merge endpoint returned status {status_code}: {response_data}")
                
                return result
                
            finally:
                # Restore original request
                flask.request = original_request
        
    except Exception as e:
        logger.exception(f"❌ Merge job failed: {str(e)}")
        raise APIError(f"Merge job failed: {str(e)}")


def update_handler(payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler for /api/v1/update endpoint - Reuses original endpoint logic
    
    Args:
        payload: Request payload containing update parameters
        
    Returns:
        Updated file result dictionary
    """
    try:
        logger.info("🔄 Starting update job processing - calling original endpoint")
        
        # Import the original endpoint handler
        from . import main
        
        # Create mock request object
        mock_request = MockRequest(payload)
        
        # Create a test request context and override the request
        with main.app.test_request_context():
            # Temporarily replace the Flask request object
            import flask
            original_request = flask.request
            flask.request = mock_request
            
            try:
                # Call the original endpoint handler
                result = main.update_excel_file()
                
                # Handle tuple response (data, status_code)
                if isinstance(result, tuple):
                    response_data, status_code = result
                    if status_code == 200:
                        return response_data
                    else:
                        raise APIError(f"Update endpoint returned status {status_code}: {response_data}")
                
                return result
                
            finally:
                # Restore original request
                flask.request = original_request
        
    except Exception as e:
        logger.exception(f"❌ Update job failed: {str(e)}")
        raise APIError(f"Update job failed: {str(e)}")


class JobHandlerRegistry:
    """Registry for mapping endpoints to handler functions"""
    
    def __init__(self):
        self._handlers: Dict[str, Callable[[Dict[str, Any]], Dict[str, Any]]] = {
            '/api/v1/extract': extract_handler,
            '/api/v1/merge': merge_handler,
            '/api/v1/update': update_handler
        }
    
    def get_handler(self, endpoint: str) -> Optional[Callable[[Dict[str, Any]], Dict[str, Any]]]:
        """Get handler function for endpoint"""
        return self._handlers.get(endpoint)
    
    def register_handler(self, endpoint: str, handler: Callable[[Dict[str, Any]], Dict[str, Any]]):
        """Register a new handler for an endpoint"""
        self._handlers[endpoint] = handler
        logger.info(f"Registered handler for endpoint: {endpoint}")
    
    def get_supported_endpoints(self) -> List[str]:
        """Get list of supported endpoints"""
        return list(self._handlers.keys())
    
    def is_supported(self, endpoint: str) -> bool:
        """Check if endpoint is supported"""
        return endpoint in self._handlers


# Global handler registry instance
handler_registry = JobHandlerRegistry()