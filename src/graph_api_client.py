"""Microsoft Graph API client for Excel range image export."""

import logging
import os
import requests
import tempfile
import time
from typing import Dict, Any, Optional, Tuple
from urllib.parse import quote

from .utils.exceptions import ExcelProcessingError
from .utils.graph_api_error_handler import (
    GraphAPIErrorHandler, 
    with_retry, 
    safe_graph_operation,
    validate_graph_response
)
from .utils.range_image_logger import range_image_logger, log_graph_api_status

logger = logging.getLogger(__name__)


class GraphAPIError(Exception):
    """Custom exception for Graph API related errors."""
    pass


class GraphAPIClient:
    """Client for Microsoft Graph API operations."""
    
    def __init__(self, client_id: str, client_secret: str, tenant_id: str,
                 max_retries: int = 3, timeout: int = 60):
        """Initialize Graph API client with Azure app credentials."""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.token_expires_at = 0
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.timeout = timeout
        self.error_handler = GraphAPIErrorHandler(max_retries=max_retries)
        
    @with_retry(max_retries=3)
    def authenticate(self) -> str:
        """Get access token using client credentials flow."""
        if self.access_token and time.time() < self.token_expires_at:
            return self.access_token
            
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }
        
        with safe_graph_operation("authentication", self.error_handler):
            response = requests.post(token_url, data=data, timeout=self.timeout)
            validate_graph_response(response, "authentication")
            
            token_data = response.json()
            self.access_token = token_data['access_token']
            # Set expiry with 5 minute buffer
            self.token_expires_at = time.time() + token_data['expires_in'] - 300
            
            log_graph_api_status(self.client_id, "authenticated", "Successfully authenticated with Microsoft Graph API")
            return self.access_token
    
    def _get_headers(self) -> Dict[str, str]:
        """Get request headers with authorization."""
        token = self.authenticate()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
    
    def upload_workbook_to_onedrive(self, file_path: str) -> str:
        """Upload Excel workbook to OneDrive and return item ID."""
        if not os.path.exists(file_path):
            raise GraphAPIError(f"File not found: {file_path}")
            
        filename = os.path.basename(file_path)
        # Use a unique filename to avoid conflicts
        unique_filename = f"temp_{int(time.time())}_{filename}"
        upload_url = f"{self.base_url}/me/drive/root:/{quote(unique_filename)}:/content"
        
        headers = {
            'Authorization': f'Bearer {self.authenticate()}',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        
        try:
            with open(file_path, 'rb') as file_data:
                response = requests.put(upload_url, headers=headers, data=file_data)
                response.raise_for_status()
                
            upload_result = response.json()
            item_id = upload_result['id']
            
            range_image_logger.info(f"📤 ONEDRIVE UPLOAD SUCCESS: {unique_filename} (ID: {item_id})")
            return item_id
            
        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to upload workbook to OneDrive: {e}")
        except KeyError as e:
            raise GraphAPIError(f"Invalid upload response format: {e}")
    
    @with_retry(max_retries=3)
    def render_range_as_image(self, item_id: str, sheet_name: str, range_str: str, 
                            width: Optional[int] = None, height: Optional[int] = None) -> bytes:
        """Render Excel range as image using Graph API."""
        # Construct the range image URL
        range_url = f"{self.base_url}/me/drive/items/{item_id}/workbook/worksheets('{sheet_name}')/range(address='{range_str}')/image"
        
        # Add optional dimensions
        params = {}
        if width:
            params['width'] = width
        if height:
            params['height'] = height
            
        headers = {
            'Authorization': f'Bearer {self.authenticate()}'
        }
        
        with safe_graph_operation(f"render_range_image:{range_str}", self.error_handler):
            response = requests.get(range_url, headers=headers, params=params, timeout=self.timeout)
            validate_graph_response(response, "render range as image")
            
            # Check if response is actually an image
            content_type = response.headers.get('content-type', '')
            if not content_type.startswith('image/'):
                raise GraphAPIError(f"Expected image response, got: {content_type}")
                
            range_image_logger.info(f"🎨 RANGE RENDER SUCCESS: {range_str} from sheet '{sheet_name}' ({len(response.content)} bytes)")
            return response.content
    
    def get_worksheet_names(self, item_id: str) -> list:
        """Get list of worksheet names from the workbook."""
        worksheets_url = f"{self.base_url}/me/drive/items/{item_id}/workbook/worksheets"
        
        try:
            response = requests.get(worksheets_url, headers=self._get_headers())
            response.raise_for_status()
            
            worksheets_data = response.json()
            worksheet_names = [ws['name'] for ws in worksheets_data.get('value', [])]
            
            logger.info(f"Found {len(worksheet_names)} worksheets: {worksheet_names}")
            return worksheet_names
            
        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to get worksheet names: {e}")
    
    def validate_range(self, item_id: str, sheet_name: str, range_str: str) -> bool:
        """Validate that a range exists and contains data."""
        range_url = f"{self.base_url}/me/drive/items/{item_id}/workbook/worksheets('{sheet_name}')/range(address='{range_str}')"
        
        try:
            response = requests.get(range_url, headers=self._get_headers())
            response.raise_for_status()
            
            range_data = response.json()
            # Check if range has any values
            values = range_data.get('values', [])
            has_data = any(any(cell for cell in row if cell is not None) for row in values)
            
            logger.info(f"Range {range_str} validation: exists=True, has_data={has_data}")
            return True
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Range validation failed for {range_str}: {e}")
            return False
    
    def cleanup_temp_file(self, item_id: str) -> None:
        """Delete the temporary file from OneDrive."""
        delete_url = f"{self.base_url}/me/drive/items/{item_id}"
        
        try:
            response = requests.delete(delete_url, headers=self._get_headers())
            response.raise_for_status()
            
            logger.info(f"Successfully deleted temporary file from OneDrive: {item_id}")
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to cleanup temporary file {item_id}: {e}")
    
    def get_range_dimensions(self, item_id: str, sheet_name: str, range_str: str) -> Tuple[int, int]:
        """Get the dimensions (rows, columns) of a range."""
        range_url = f"{self.base_url}/me/drive/items/{item_id}/workbook/worksheets('{sheet_name}')/range(address='{range_str}')"
        
        try:
            response = requests.get(range_url, headers=self._get_headers())
            response.raise_for_status()
            
            range_data = response.json()
            row_count = range_data.get('rowCount', 0)
            column_count = range_data.get('columnCount', 0)
            
            logger.info(f"Range {range_str} dimensions: {row_count} rows x {column_count} columns")
            return row_count, column_count
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to get range dimensions for {range_str}: {e}")
            return 0, 0