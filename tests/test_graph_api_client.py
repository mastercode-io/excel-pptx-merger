"""Unit tests for Graph API client."""

import pytest
import unittest.mock as mock
from unittest.mock import Mock, patch, MagicMock
import requests
import time

from src.graph_api_client import GraphAPIClient, GraphAPIError
from src.utils.graph_api_error_handler import GraphAPIRetryableError, GraphAPIFatalError
from src.utils.exceptions import ExcelProcessingError


class TestGraphAPIClient:
    """Test cases for GraphAPIClient."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.client = GraphAPIClient(
            client_id="test_client_id",
            client_secret="test_client_secret", 
            tenant_id="test_tenant_id",
            sharepoint_config={
                "enabled": True,
                "site_id": "test_site_id",
                "drive_id": "test_drive_id",
                "temp_folder_path": "/Temp/ExcelProcessing"
            }
        )
    
    def test_init(self):
        """Test client initialization."""
        assert self.client.client_id == "test_client_id"
        assert self.client.client_secret == "test_client_secret"
        assert self.client.tenant_id == "test_tenant_id"
        assert self.client.access_token is None
        assert self.client.token_expires_at == 0
        assert self.client.timeout == 60
    
    @patch('src.graph_api_client.requests.post')
    def test_authenticate_success(self, mock_post):
        """Test successful authentication."""
        # Mock successful token response
        mock_response = Mock()
        mock_response.json.return_value = {
            'access_token': 'test_token',
            'expires_in': 3600
        }
        mock_response.raise_for_status.return_value = None
        mock_post.return_value = mock_response
        
        token = self.client.authenticate()
        
        assert token == 'test_token'
        assert self.client.access_token == 'test_token'
        assert self.client.token_expires_at > time.time()
        
        # Verify token endpoint was called correctly
        mock_post.assert_called_once()
        call_args = mock_post.call_args
        assert f"/{self.client.tenant_id}/oauth2/v2.0/token" in call_args[0][0]
    
    @patch('src.graph_api_client.requests.post')
    def test_authenticate_cached_token(self, mock_post):
        """Test authentication with cached valid token."""
        # Set up cached token
        self.client.access_token = 'cached_token'
        self.client.token_expires_at = time.time() + 1800  # 30 minutes from now
        
        token = self.client.authenticate()
        
        assert token == 'cached_token'
        # Should not make HTTP request
        mock_post.assert_not_called()
    
    @patch('src.graph_api_client.requests.post')
    def test_authenticate_http_error(self, mock_post):
        """Test authentication with HTTP error."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError("401 Unauthorized")
        mock_post.return_value = mock_response
        
        with pytest.raises(Exception):  # Will be wrapped by retry decorator
            self.client.authenticate()
    
    def test_get_headers(self):
        """Test header generation."""
        with patch.object(self.client, 'authenticate', return_value='test_token'):
            headers = self.client._get_headers()
            
            expected_headers = {
                'Authorization': 'Bearer test_token',
                'Content-Type': 'application/json'
            }
            assert headers == expected_headers
    
    @patch('src.graph_api_client.requests.put')
    @patch('builtins.open', mock.mock_open(read_data=b'fake_excel_data'))
    @patch('os.path.exists', return_value=True)
    def test_upload_workbook_success(self, mock_exists, mock_put):
        """Test successful workbook upload."""
        # Mock successful upload response
        mock_response = Mock()
        mock_response.json.return_value = {'id': 'test_item_id'}
        mock_response.raise_for_status.return_value = None
        mock_put.return_value = mock_response
        
        with patch.object(self.client, 'authenticate', return_value='test_token'):
            item_id = self.client.upload_workbook_to_onedrive('/path/to/test.xlsx')
            
            assert item_id == 'test_item_id'
            mock_put.assert_called_once()
    
    @patch('os.path.exists', return_value=False)
    def test_upload_workbook_file_not_found(self, mock_exists):
        """Test workbook upload with missing file."""
        with pytest.raises(GraphAPIError, match="File not found"):
            self.client.upload_workbook_to_onedrive('/path/to/missing.xlsx')
    
    @patch('src.graph_api_client.requests.get')
    def test_render_range_as_image_success(self, mock_get):
        """Test successful range image rendering."""
        # Mock successful image response
        mock_response = Mock()
        mock_response.content = b'fake_image_data'
        mock_response.headers = {'content-type': 'image/png'}
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, 'authenticate', return_value='test_token'):
            image_data = self.client.render_range_as_image(
                item_id='test_item_id',
                sheet_name='Sheet1', 
                range_str='A1:C10'
            )
            
            assert image_data == b'fake_image_data'
            mock_get.assert_called_once()
    
    @patch('src.graph_api_client.requests.get')
    def test_render_range_as_image_invalid_content_type(self, mock_get):
        """Test range image rendering with invalid content type."""
        # Mock response with non-image content type
        mock_response = Mock()
        mock_response.content = b'not_image_data'
        mock_response.headers = {'content-type': 'text/html'}
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, 'authenticate', return_value='test_token'):
            with pytest.raises(ExcelProcessingError, match="Graph API operation failed with non-retryable error"):
                self.client.render_range_as_image(
                    item_id='test_item_id',
                    sheet_name='Sheet1',
                    range_str='A1:C10'
                )
    
    @patch('src.graph_api_client.requests.get')
    def test_get_worksheet_names_success(self, mock_get):
        """Test successful worksheet names retrieval."""
        # Mock successful worksheets response
        mock_response = Mock()
        mock_response.json.return_value = {
            'value': [
                {'name': 'Sheet1'},
                {'name': 'Sheet2'},
                {'name': 'Data'}
            ]
        }
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            worksheets = self.client.get_worksheet_names('test_item_id')
            
            assert worksheets == ['Sheet1', 'Sheet2', 'Data']
            mock_get.assert_called_once()
    
    @patch('src.graph_api_client.requests.get')
    def test_validate_range_success(self, mock_get):
        """Test successful range validation."""
        # Mock successful range response with data
        mock_response = Mock()
        mock_response.json.return_value = {
            'values': [
                ['Header1', 'Header2'],
                ['Data1', 'Data2'],
                ['Data3', None]
            ]
        }
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            is_valid = self.client.validate_range('test_item_id', 'Sheet1', 'A1:B3')
            
            assert is_valid is True
            mock_get.assert_called_once()
    
    @patch('src.graph_api_client.requests.get')
    def test_validate_range_empty(self, mock_get):
        """Test range validation with empty range."""
        # Mock range response with no data
        mock_response = Mock()
        mock_response.json.return_value = {
            'values': [
                [None, None],
                [None, None]
            ]
        }
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            is_valid = self.client.validate_range('test_item_id', 'Sheet1', 'A1:B2')
            
            assert is_valid is True  # Range exists but has no data
    
    @patch('src.graph_api_client.requests.get')
    def test_validate_range_error(self, mock_get):
        """Test range validation with API error."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError("404 Not Found")
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            is_valid = self.client.validate_range('test_item_id', 'Sheet1', 'InvalidRange')
            
            assert is_valid is False
    
    @patch('src.graph_api_client.requests.delete')
    def test_cleanup_temp_file_success(self, mock_delete):
        """Test successful temp file cleanup."""
        mock_response = Mock()
        mock_response.raise_for_status.return_value = None
        mock_delete.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            # Should not raise exception
            self.client.cleanup_temp_file('test_item_id')
            mock_delete.assert_called_once()
    
    @patch('src.graph_api_client.requests.delete')
    def test_cleanup_temp_file_error(self, mock_delete):
        """Test temp file cleanup with error (should not raise)."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError("404 Not Found")
        mock_delete.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            # Should not raise exception, just log warning
            self.client.cleanup_temp_file('test_item_id')
            mock_delete.assert_called_once()
    
    @patch('src.graph_api_client.requests.get')
    def test_get_range_dimensions_success(self, mock_get):
        """Test successful range dimensions retrieval."""
        mock_response = Mock()
        mock_response.json.return_value = {
            'rowCount': 5,
            'columnCount': 3
        }
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            rows, cols = self.client.get_range_dimensions('test_item_id', 'Sheet1', 'A1:C5')
            
            assert rows == 5
            assert cols == 3
            mock_get.assert_called_once()
    
    @patch('src.graph_api_client.requests.get')
    def test_get_range_dimensions_error(self, mock_get):
        """Test range dimensions with API error."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = requests.exceptions.HTTPError("400 Bad Request")
        mock_get.return_value = mock_response
        
        with patch.object(self.client, '_get_headers', return_value={'Authorization': 'Bearer test_token'}):
            rows, cols = self.client.get_range_dimensions('test_item_id', 'Sheet1', 'InvalidRange')
            
            assert rows == 0
            assert cols == 0