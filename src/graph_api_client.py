"""Microsoft Graph API client for Excel range image export."""

import base64
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
    validate_graph_response,
)
from .utils.range_image_logger import range_image_logger, log_graph_api_status

logger = logging.getLogger(__name__)


class GraphAPIError(Exception):
    """Custom exception for Graph API related errors."""

    pass


class GraphAPIClient:
    """Client for Microsoft Graph API operations."""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        tenant_id: str,
        sharepoint_config: Optional[Dict[str, Any]] = None,
        max_retries: int = 3,
        timeout: int = 60,
    ):
        """Initialize Graph API client with Azure app credentials and SharePoint config."""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.sharepoint_config = sharepoint_config or {}
        self.access_token = None
        self.token_expires_at = 0
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.timeout = timeout
        self.error_handler = GraphAPIErrorHandler(max_retries=max_retries)
        
        # SharePoint context for client credentials flow
        self.current_site_id = None
        self.current_drive_id = None

    @with_retry(max_retries=3)
    def authenticate(self) -> str:
        """Get access token using client credentials flow."""
        if self.access_token and time.time() < self.token_expires_at:
            return self.access_token

        token_url = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        )

        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }

        with safe_graph_operation("authentication", self.error_handler):
            logger.info(f"🔐 Authenticating with tenant: {self.tenant_id}")
            logger.info(f"🔐 Client ID: {self.client_id[:8]}...")
            
            response = requests.post(token_url, data=data, timeout=self.timeout)
            
            # Enhanced authentication error debugging
            if not response.ok:
                logger.error(f"❌ Authentication failed:")
                logger.error(f"   Status: {response.status_code}")
                logger.error(f"   URL: {token_url}")
                logger.error(f"   Tenant ID: {self.tenant_id}")
                logger.error(f"   Client ID: {self.client_id[:8]}...")
                
                try:
                    error_details = response.json()
                    logger.error(f"   Error details: {error_details}")
                    
                    error_code = error_details.get("error", "")
                    error_desc = error_details.get("error_description", "")
                    
                    if error_code == "invalid_client":
                        logger.error("💡 Invalid client credentials - check client_id and client_secret")
                    elif error_code == "invalid_request":
                        if "AADSTS90002" in error_desc:
                            logger.error("💡 TENANT NOT FOUND - The tenant_id is invalid or tenant has no active subscriptions")
                            logger.error("💡 Solution: Check with your SharePoint admin for the correct tenant_id")
                        else:
                            logger.error("💡 Invalid request parameters")
                    elif error_code == "unauthorized_client":
                        if "AADSTS700016" in error_desc:
                            logger.error("💡 APPLICATION NOT FOUND IN TENANT - The Azure app is not registered in this tenant")
                            logger.error("💡 Solution: Register the application in the correct tenant or use the right tenant_id/client_id pair")
                        else:
                            logger.error("💡 Client not authorized - check app registration permissions")
                    elif "tenant" in error_desc.lower():
                        logger.error("💡 Tenant ID issue - verify the correct tenant_id is being used")
                    
                except Exception as parse_error:
                    logger.error(f"   Raw response: {response.text}")
                    logger.error(f"   Parse error: {parse_error}")
            
            validate_graph_response(response, "authentication")

            token_data = response.json()
            self.access_token = token_data["access_token"]
            # Set expiry with 5 minute buffer
            self.token_expires_at = time.time() + token_data["expires_in"] - 300

            log_graph_api_status(
                self.client_id,
                "authenticated",
                "Successfully authenticated with Microsoft Graph API",
            )
            return self.access_token

    def _get_headers(self) -> Dict[str, str]:
        """Get request headers with authorization."""
        token = self.authenticate()
        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    def _build_item_url(self, item_id: str, endpoint_path: str = "") -> str:
        """Build appropriate Graph API URL for an item based on available context.
        
        Args:
            item_id: The drive item ID
            endpoint_path: Additional path after the item (e.g., "/workbook/worksheets")
            
        Returns:
            Properly formatted Graph API URL
        """
        if self.current_site_id and self.current_drive_id:
            # Use site-specific endpoint
            base_url = f"{self.base_url}/sites/{self.current_site_id}/drives/{self.current_drive_id}/items/{item_id}"
        elif self.current_drive_id:
            # Use drive-specific endpoint
            base_url = f"{self.base_url}/drives/{self.current_drive_id}/items/{item_id}"
        else:
            # Fallback to /me endpoint (will likely fail with client credentials)
            logger.warning("No SharePoint context available, falling back to /me endpoint")
            base_url = f"{self.base_url}/me/drive/items/{item_id}"
        
        return f"{base_url}{endpoint_path}"

    def _get_sharepoint_upload_url(self, filename: str) -> Optional[str]:
        """Get SharePoint upload URL based on configuration."""
        if not self.sharepoint_config.get("enabled"):
            return None

        site_id = self.sharepoint_config.get("site_id")
        drive_id = self.sharepoint_config.get("drive_id")
        temp_folder = self.sharepoint_config.get(
            "temp_folder_path", "/Temp/ExcelProcessing"
        )

        if not site_id:
            raise GraphAPIError(
                "SharePoint site_id is required when SharePoint is enabled"
            )

        # Ensure temp folder starts with /
        if not temp_folder.startswith("/"):
            temp_folder = "/" + temp_folder

        # Create unique filename to avoid conflicts
        unique_filename = f"temp_{int(time.time())}_{filename}"
        file_path = f"{temp_folder}/{unique_filename}"

        if drive_id:
            # Specific drive in SharePoint site
            return f"{self.base_url}/sites/{quote(site_id)}/drives/{quote(drive_id)}/root:{quote(file_path)}:/content"
        else:
            # Default drive in SharePoint site
            return f"{self.base_url}/sites/{quote(site_id)}/drive/root:{quote(file_path)}:/content"

    def _encode_sharing_url(self, url: str) -> str:
        """Encode a SharePoint URL for use with the Shares API.

        Args:
            url: SharePoint URL to encode

        Returns:
            Encoded share token for use with Shares API
        """
        # Encode URL to base64
        url_bytes = url.encode("utf-8")
        base64_encoded = base64.b64encode(url_bytes).decode("utf-8")

        # Convert to unpadded base64url format
        # Remove = padding, replace / with _, replace + with -
        share_token = base64_encoded.rstrip("=").replace("/", "_").replace("+", "-")

        # Add u! prefix
        return f"u!{share_token}"

    def _get_driveitem_from_share_token(self, share_token: str) -> Dict[str, str]:
        """Get drive item information from a share token.
        
        Args:
            share_token: The share token extracted from sharing link
            
        Returns:
            Dictionary containing item_id, drive_id, and optionally site_id
        """
        try:
            # Encode the share token for the API (add u! prefix)
            encoded_token = f"u!{share_token}"
            
            # Use the shares endpoint to get the driveItem
            shares_url = f"{self.base_url}/shares/{encoded_token}/driveItem"
            
            headers = {"Authorization": f"Bearer {self.authenticate()}"}
            
            response = requests.get(shares_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()
            
            drive_item = response.json()
            item_id = drive_item["id"]
            
            # Extract parent reference information
            parent_ref = drive_item.get("parentReference", {})
            drive_id = parent_ref.get("driveId")
            site_id = parent_ref.get("siteId")
            
            # Store context for subsequent API calls
            self.current_drive_id = drive_id
            self.current_site_id = site_id
            
            result = {
                "item_id": item_id,
                "drive_id": drive_id,
                "site_id": site_id
            }
            
            logger.info(f"Resolved share token via Shares API:")
            logger.info(f"  Item ID: {item_id}")
            logger.info(f"  Drive ID: {drive_id}")
            logger.info(f"  Site ID: {site_id}")
            
            return result
            
        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to resolve share token via Shares API: {e}")
        except KeyError as e:
            raise GraphAPIError(f"Invalid response from Shares API: {e}")

    def get_driveitem_from_sharing_url(self, sharepoint_url: str) -> Dict[str, str]:
        """Get drive item information from SharePoint URL using Shares API.

        Args:
            sharepoint_url: SharePoint URL (any format)

        Returns:
            Dictionary containing item_id, drive_id, and optionally site_id
        """
        try:
            # Encode the URL for the Shares API
            share_token = self._encode_sharing_url(sharepoint_url)

            # Use the shares endpoint to get the driveItem
            shares_url = f"{self.base_url}/shares/{share_token}/driveItem"

            headers = {"Authorization": f"Bearer {self.authenticate()}"}

            response = requests.get(shares_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            drive_item = response.json()
            item_id = drive_item["id"]
            
            # Extract parent reference information
            parent_ref = drive_item.get("parentReference", {})
            drive_id = parent_ref.get("driveId")
            site_id = parent_ref.get("siteId")
            
            # Store context for subsequent API calls
            self.current_drive_id = drive_id
            self.current_site_id = site_id
            
            result = {
                "item_id": item_id,
                "drive_id": drive_id,
            }
            
            if site_id:
                result["site_id"] = site_id
            
            logger.info(f"Resolved SharePoint URL via Shares API:")
            logger.info(f"  Item ID: {item_id}")
            logger.info(f"  Drive ID: {drive_id}")
            logger.info(f"  Site ID: {site_id}")
            
            return result

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to resolve SharePoint URL via Shares API: {e}")
        except KeyError as e:
            raise GraphAPIError(f"Invalid response from Shares API: {e}")

    def get_sharepoint_item_id_from_url(self, sharepoint_url: str) -> str:
        """Extract item ID from SharePoint URL using Graph API resolution."""
        try:
            logger.info(f"🔗 Resolving SharePoint URL: {sharepoint_url}")
            from .utils.sharepoint_url_parser import SharePointUrlParser

            parser = SharePointUrlParser()
            parsed = parser.parse_sharepoint_url(sharepoint_url)

            if not parsed:
                raise GraphAPIError(f"Invalid SharePoint URL format: {sharepoint_url}")

            logger.info(f"📝 Parsed URL components: {parsed}")

            # Check if this URL requires the Shares API (Doc.aspx format or sharing link)
            if parsed.get("requires_shares_api", False):
                logger.info(f"🌐 Using Shares API for URL type: {parsed.get('url_type')}")

                # For all Shares API URLs (sharing links and Doc.aspx), encode the full URL
                # The Microsoft Graph Shares API requires the full URL to be base64-encoded
                drive_item_info = self.get_driveitem_from_sharing_url(sharepoint_url)

                item_id = drive_item_info["item_id"]
                logger.info(f"✅ Shares API resolved to item ID: {item_id}")
                return item_id

            # Use traditional path-based resolution for standard URLs
            logger.info(f"🔧 Using path-based resolution for URL: {sharepoint_url}")

            # Get site ID from site name
            site_id = self._resolve_site_id(parsed["hostname"], parsed["site_name"])

            # Get drive ID from library name
            drive_id = self._resolve_drive_id(site_id, parsed["normalized_library"])

            # Get item ID from file path
            item_id = self._resolve_item_id(site_id, drive_id, parsed["file_path"])

            logger.info(f"✅ Path-based resolution completed - item ID: {item_id}")
            return item_id

        except Exception as e:
            logger.error(f"❌ Failed to resolve SharePoint URL: {sharepoint_url}")
            logger.error(f"   Error: {e}")
            raise GraphAPIError(f"Failed to resolve item ID from SharePoint URL: {e}")

    def _resolve_site_id(self, hostname: str, site_name: str) -> str:
        """Resolve site name to site ID using Graph API."""
        site_url = f"{self.base_url}/sites/{hostname}:/sites/{site_name}"

        headers = {"Authorization": f"Bearer {self.authenticate()}"}

        try:
            response = requests.get(site_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            site_data = response.json()
            site_id = site_data["id"]

            logger.info(f"Resolved site '{site_name}' to ID: {site_id}")
            return site_id

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to resolve site ID for '{site_name}': {e}")

    def _resolve_drive_id(self, site_id: str, library_name: str) -> str:
        """Resolve library name to drive ID using Graph API."""
        drives_url = f"{self.base_url}/sites/{site_id}/drives"

        headers = {"Authorization": f"Bearer {self.authenticate()}"}

        try:
            response = requests.get(drives_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            drives_data = response.json()

            # Find drive by name
            for drive in drives_data.get("value", []):
                if drive.get("name") == library_name:
                    drive_id = drive["id"]
                    logger.info(
                        f"Resolved library '{library_name}' to drive ID: {drive_id}"
                    )
                    return drive_id

            # If not found by name, try to find default document library
            for drive in drives_data.get("value", []):
                if drive.get("driveType") == "documentLibrary":
                    drive_id = drive["id"]
                    logger.info(
                        f"Using default document library as drive ID: {drive_id}"
                    )
                    return drive_id

            raise GraphAPIError(f"Could not find drive for library '{library_name}'")

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(
                f"Failed to resolve drive ID for library '{library_name}': {e}"
            )

    def _resolve_item_id(self, site_id: str, drive_id: str, file_path: str) -> str:
        """Resolve file path to item ID using Graph API."""
        # URL encode the file path
        encoded_path = quote(file_path)
        item_url = (
            f"{self.base_url}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"
        )

        headers = {"Authorization": f"Bearer {self.authenticate()}"}

        try:
            response = requests.get(item_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            item_data = response.json()
            item_id = item_data["id"]

            logger.info(f"Resolved file path '{file_path}' to item ID: {item_id}")
            return item_id

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(
                f"Failed to resolve item ID for file '{file_path}': {e}"
            )

    def download_file_from_sharing_url(self, sharepoint_url: str) -> bytes:
        """Download file directly from SharePoint URL using Shares API.

        Args:
            sharepoint_url: SharePoint URL (any format)

        Returns:
            File content as bytes
        """
        try:
            # Encode the URL for the Shares API
            share_token = self._encode_sharing_url(sharepoint_url)

            # Use the shares endpoint to download the content
            download_url = f"{self.base_url}/shares/{share_token}/driveItem/content"

            headers = {"Authorization": f"Bearer {self.authenticate()}"}

            response = requests.get(download_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            logger.info(
                f"Downloaded file from SharePoint via Shares API: {len(response.content)} bytes"
            )
            return response.content

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to download file via Shares API: {e}")

    def download_file_to_memory(self, item_id: str) -> bytes:
        """Download file from SharePoint/OneDrive to memory using item ID."""
        download_url = self._build_item_url(item_id, "/content")

        headers = {"Authorization": f"Bearer {self.authenticate()}"}

        try:
            response = requests.get(download_url, headers=headers, timeout=self.timeout)
            response.raise_for_status()

            logger.info(
                f"Downloaded file from SharePoint/OneDrive: {len(response.content)} bytes"
            )
            return response.content

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(
                f"Failed to download file from SharePoint/OneDrive: {e}"
            )

    def validate_sharepoint_config(self) -> bool:
        """Validate SharePoint configuration."""
        if not self.sharepoint_config.get("enabled"):
            return False

        # For URL-based operations, we don't need pre-configured site_id
        # The validation will happen during URL resolution
        return True

    def validate_sharepoint_config_for_upload(self) -> bool:
        """Validate SharePoint configuration for file upload operations."""
        if not self.sharepoint_config.get("enabled"):
            return False

        # For uploads, we need explicit site_id configuration
        site_id = self.sharepoint_config.get("site_id")
        if not site_id:
            logger.error(
                "SharePoint enabled but site_id not provided (required for uploads)"
            )
            return False

        return True

    def upload_workbook_to_sharepoint(self, file_path: str) -> str:
        """Upload Excel workbook to SharePoint and return item ID."""
        if not os.path.exists(file_path):
            raise GraphAPIError(f"File not found: {file_path}")

        # Validate SharePoint configuration for upload operations
        if not self.validate_sharepoint_config_for_upload():
            raise GraphAPIError(
                "SharePoint configuration with site_id is required for file uploads"
            )

        filename = os.path.basename(file_path)
        upload_url = self._get_sharepoint_upload_url(filename)

        if not upload_url:
            raise GraphAPIError("Failed to generate SharePoint upload URL")

        headers = {
            "Authorization": f"Bearer {self.authenticate()}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }

        try:
            with open(file_path, "rb") as file_data:
                response = requests.put(upload_url, headers=headers, data=file_data)
                response.raise_for_status()

            upload_result = response.json()
            item_id = upload_result["id"]

            range_image_logger.info(
                f"📤 SHAREPOINT UPLOAD SUCCESS: {filename} (ID: {item_id})"
            )
            return item_id

        except requests.exceptions.RequestException as e:
            raise GraphAPIError(f"Failed to upload workbook to SharePoint: {e}")
        except KeyError as e:
            raise GraphAPIError(f"Invalid upload response format: {e}")

    # Keep backward compatibility
    def upload_workbook_to_onedrive(self, file_path: str) -> str:
        """Legacy method - now redirects to SharePoint upload."""
        return self.upload_workbook_to_sharepoint(file_path)

    @with_retry(max_retries=3)
    def render_range_as_image(
        self,
        item_id: str,
        sheet_name: str,
        range_str: str,
        width: Optional[int] = None,
        height: Optional[int] = None,
    ) -> bytes:
        """Render Excel range as image using Graph API."""
        # Construct the range image URL
        range_url = self._build_item_url(item_id, f"/workbook/worksheets('{sheet_name}')/range(address='{range_str}')/image")

        # Add optional dimensions
        params = {}
        if width:
            params["width"] = width
        if height:
            params["height"] = height

        headers = {"Authorization": f"Bearer {self.authenticate()}"}

        with safe_graph_operation(
            f"render_range_image:{range_str}", self.error_handler
        ):
            response = requests.get(
                range_url, headers=headers, params=params, timeout=self.timeout
            )
            
            # Enhanced error logging for debugging range rendering issues
            if not response.ok:
                logger.error(f"❌ Range rendering failed:")
                logger.error(f"   Status: {response.status_code}")
                logger.error(f"   URL: {range_url}")
                try:
                    error_data = response.json()
                    error_code = error_data.get("error", {}).get("code", "")
                    error_msg = error_data.get("error", {}).get("message", "")
                    logger.error(f"   Error code: {error_code}")
                    logger.error(f"   Error message: {error_msg}")
                    
                    # Check for specific workbook-related errors
                    if "InvalidSession" in error_code:
                        logger.error("💡 Workbook session required - need to create session first")
                    elif "ItemNotFound" in error_code:
                        logger.error(f"💡 Sheet '{sheet_name}' or range '{range_str}' not found")
                    elif "InvalidRequest" in error_code:
                        logger.error(f"💡 Invalid range format or request: {range_str}")
                    elif "AccessDenied" in error_code:
                        logger.error("💡 Insufficient permissions for workbook operations")
                    elif "BadRequest" in error_code and "session" in error_msg.lower():
                        logger.error("💡 Workbook session issue - may need to create session")
                        
                except Exception as e:
                    logger.error(f"   Raw response: {response.text[:500]}")
                    logger.error(f"   Parse error: {e}")
            
            validate_graph_response(response, "render range as image")

            # Check if response is actually an image
            content_type = response.headers.get("content-type", "")
            if not content_type.startswith("image/"):
                raise GraphAPIError(f"Expected image response, got: {content_type}")

            range_image_logger.info(
                f"🎨 RANGE RENDER SUCCESS: {range_str} from sheet '{sheet_name}' ({len(response.content)} bytes)"
            )
            return response.content

    @with_retry(max_retries=3)
    def get_worksheet_names(self, item_id: str) -> list:
        """Get list of worksheet names from the workbook."""
        try:
            # First, validate the file format and get file info
            logger.info(f"🔍 Validating workbook format for item: {item_id}")
            file_info = self._get_file_info(item_id)
            logger.info(f"📄 File info: {file_info}")
            
            # Check if it's a valid Excel workbook
            file_name = file_info.get("name", "unknown")
            mime_type = file_info.get("file", {}).get("mimeType", "unknown")
            
            if not file_name.lower().endswith(('.xlsx', '.xlsm')):
                logger.warning(f"⚠️ File might not be a modern Excel format: {file_name}")
                logger.warning(f"   MIME type: {mime_type}")
            
            worksheets_url = self._build_item_url(item_id, "/workbook/worksheets")
            logger.info(f"🔍 Attempting to get worksheets from: {worksheets_url}")
            logger.info(f"🔧 Using SharePoint context - Site ID: {self.current_site_id}, Drive ID: {self.current_drive_id}")
            
            response = requests.get(worksheets_url, headers=self._get_headers())
            
            # Enhanced error logging for debugging
            logger.info(f"📡 Response status: {response.status_code}")
            logger.info(f"📡 Response headers: {dict(response.headers)}")
            
            if not response.ok:
                logger.error(f"❌ Graph API Error Details:")
                logger.error(f"   Status: {response.status_code}")
                logger.error(f"   Reason: {response.reason}")
                logger.error(f"   URL: {worksheets_url}")
                logger.error(f"   File: {file_name}")
                logger.error(f"   MIME: {mime_type}")
                try:
                    error_details = response.json()
                    logger.error(f"   Error Response: {error_details}")
                    
                    # Check for specific error types
                    error_code = error_details.get("error", {}).get("code", "")
                    if error_code == "InvalidRequest":
                        logger.error("💡 This may indicate the file is not a valid Excel workbook or is corrupted")
                    elif error_code == "Forbidden":
                        logger.error("💡 This may indicate insufficient permissions to access the workbook")
                    elif error_code == "NotFound":
                        logger.error("💡 The file may have been moved or deleted")
                        
                except Exception as parse_error:
                    logger.error(f"   Raw Response: {response.text[:500]}")
                    logger.error(f"   Failed to parse error: {parse_error}")
            
            response.raise_for_status()

            worksheets_data = response.json()
            worksheet_names = [ws["name"] for ws in worksheets_data.get("value", [])]

            logger.info(f"✅ Found {len(worksheet_names)} worksheets: {worksheet_names}")
            return worksheet_names

        except requests.exceptions.RequestException as e:
            logger.error(f"❌ Graph API Request failed: {e}")
            logger.error(f"   Item ID: {item_id}")
            raise GraphAPIError(f"Failed to get worksheet names: {e}")
    
    def _get_file_info(self, item_id: str) -> dict:
        """Get file information including name and MIME type."""
        file_info_url = self._build_item_url(item_id)
        
        try:
            response = requests.get(file_info_url, headers=self._get_headers())
            response.raise_for_status()
            
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to get file info: {e}")
            return {}

    def validate_range(self, item_id: str, sheet_name: str, range_str: str) -> bool:
        """Validate that a range exists and contains data."""
        range_url = self._build_item_url(item_id, f"/workbook/worksheets('{sheet_name}')/range(address='{range_str}')")

        try:
            response = requests.get(range_url, headers=self._get_headers())
            response.raise_for_status()

            range_data = response.json()
            # Check if range has any values
            values = range_data.get("values", [])
            has_data = any(
                any(cell for cell in row if cell is not None) for row in values
            )

            logger.info(
                f"Range {range_str} validation: exists=True, has_data={has_data}"
            )
            return True

        except requests.exceptions.RequestException as e:
            logger.warning(f"Range validation failed for {range_str}: {e}")
            return False

    def cleanup_temp_file(self, item_id: str) -> None:
        """Delete the temporary file from OneDrive."""
        delete_url = self._build_item_url(item_id)

        try:
            response = requests.delete(delete_url, headers=self._get_headers())
            response.raise_for_status()

            logger.info(f"Successfully deleted temporary file from OneDrive: {item_id}")

        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to cleanup temporary file {item_id}: {e}")

    def get_range_dimensions(
        self, item_id: str, sheet_name: str, range_str: str
    ) -> Tuple[int, int]:
        """Get the dimensions (rows, columns) of a range."""
        range_url = self._build_item_url(item_id, f"/workbook/worksheets('{sheet_name}')/range(address='{range_str}')")

        try:
            response = requests.get(range_url, headers=self._get_headers())
            response.raise_for_status()

            range_data = response.json()
            row_count = range_data.get("rowCount", 0)
            column_count = range_data.get("columnCount", 0)

            logger.info(
                f"Range {range_str} dimensions: {row_count} rows x {column_count} columns"
            )
            return row_count, column_count

        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to get range dimensions for {range_str}: {e}")
            return 0, 0
