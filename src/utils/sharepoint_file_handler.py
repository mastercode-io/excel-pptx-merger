"""Centralized SharePoint file handling for all endpoints."""

import io
import logging
from typing import Dict, Any, Optional, Tuple

from ..graph_api_config import get_graph_api_credentials
from .exceptions import ValidationError, AuthenticationError

logger = logging.getLogger(__name__)


class SharePointFileHandler:
    """Centralized handler for SharePoint file operations across all endpoints."""
    
    def __init__(self, sharepoint_config: Dict[str, Any]):
        """Initialize SharePoint file handler.
        
        Args:
            sharepoint_config: SharePoint configuration from global_settings
        """
        self.sharepoint_config = sharepoint_config
        self.graph_credentials = None
        self._load_credentials()
    
    def _load_credentials(self) -> None:
        """Load Graph API credentials with detailed diagnostic logging."""
        logger.info("🔧 Starting Graph API credential loading process...")
        
        try:
            # Log SharePoint configuration details
            logger.info(f"📋 SharePoint config received: {dict(self.sharepoint_config)}")
            config_tenant_id = self.sharepoint_config.get("tenant_id")
            logger.info(f"📋 Extracted tenant_id from config: {config_tenant_id}")
            
            # Log environment variable status (masked for security)
            import os
            env_client_id = os.getenv("GRAPH_CLIENT_ID", "")
            env_client_secret = os.getenv("GRAPH_CLIENT_SECRET", "")
            env_tenant_id = os.getenv("GRAPH_TENANT_ID", "")
            
            logger.info(f"🔑 Environment variables status:")
            logger.info(f"🔑   GRAPH_CLIENT_ID: {'SET (' + env_client_id[:8] + '...)' if env_client_id else 'NOT SET'}")
            logger.info(f"🔑   GRAPH_CLIENT_SECRET: {'SET (' + env_client_secret[:4] + '...)' if env_client_secret else 'NOT SET'}")
            logger.info(f"🔑   GRAPH_TENANT_ID: {'SET (' + env_tenant_id[:8] + '...)' if env_tenant_id else 'NOT SET'}")
            
            # Check for graph_api.env file
            possible_paths = ["config/graph_api.env", "graph_api.env", ".env"]
            for path in possible_paths:
                if os.path.exists(path):
                    logger.info(f"📁 Found Graph API config file: {path}")
                    break
            else:
                logger.warning(f"📁 No Graph API config file found in: {possible_paths}")
            
            # Log current working directory
            logger.info(f"📂 Current working directory: {os.getcwd()}")
            
            # Call get_graph_api_credentials with detailed logging
            logger.info(f"🔧 Calling get_graph_api_credentials(tenant_id='{config_tenant_id}')...")
            self.graph_credentials = get_graph_api_credentials(config_tenant_id)
            logger.info(f"🔧 get_graph_api_credentials() returned: {type(self.graph_credentials)} - {bool(self.graph_credentials)}")
            
            if self.graph_credentials:
                logger.info("✅ Graph API credentials loaded successfully - SharePoint access enabled")
                logger.info(f"✅ Credential keys available: {list(self.graph_credentials.keys())}")
                logger.debug(
                    f"✅ Loaded credentials: client_id={self.graph_credentials.get('client_id', '')[:8]}..."
                )
            else:
                logger.error("❌ Graph API credentials are None - SharePoint access disabled")
                logger.error("❌ This will cause SharePoint file access to fail!")
                logger.error("❌ Check environment variables: GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID")
                logger.error("❌ Or ensure config/graph_api.env file exists and is readable")
                
        except Exception as e:
            logger.error(f"❌ Exception during Graph API credential loading: {e}")
            logger.error(f"❌ Exception type: {type(e)}")
            import traceback
            logger.error(f"❌ Traceback: {traceback.format_exc()}")
            self.graph_credentials = None
    
    def is_configured(self) -> bool:
        """Check if SharePoint is properly configured for file access.
        
        Returns:
            True if SharePoint is enabled and credentials are available
        """
        sharepoint_enabled = self.sharepoint_config.get("enabled", False)
        credentials_available = bool(self.graph_credentials)
        
        logger.debug(f"SharePoint configuration check - enabled: {sharepoint_enabled}, credentials: {credentials_available}")
        
        return sharepoint_enabled and credentials_available
    
    def validate_sharepoint_access(self) -> None:
        """Validate SharePoint access requirements.
        
        Raises:
            ValidationError: If SharePoint is not properly configured
            AuthenticationError: If Graph API credentials are missing
        """
        if not self.sharepoint_config.get("enabled"):
            raise ValidationError(
                "SharePoint must be enabled in config.global_settings.sharepoint"
            )
        
        if not self.graph_credentials:
            raise AuthenticationError(
                "Graph API credentials required for SharePoint file access"
            )
    
    def download_file(
        self, 
        sharepoint_url: Optional[str] = None, 
        sharepoint_item_id: Optional[str] = None,
        default_filename: str = "sharepoint_file.xlsx"
    ) -> io.BytesIO:
        """Download file from SharePoint using URL or item ID.
        
        Args:
            sharepoint_url: SharePoint sharing URL
            sharepoint_item_id: Direct SharePoint item ID
            default_filename: Default filename to assign to downloaded file
            
        Returns:
            BytesIO object containing the downloaded file
            
        Raises:
            ValidationError: If neither URL nor item ID provided, or download fails
            AuthenticationError: If Graph API credentials are missing
        """
        # Validate input parameters
        if not sharepoint_url and not sharepoint_item_id:
            raise ValidationError(
                "Either sharepoint_url or sharepoint_item_id must be provided"
            )
        
        # Validate SharePoint access
        self.validate_sharepoint_access()
        
        try:
            # Initialize Graph API client
            from ..graph_api_client import GraphAPIClient
            
            logger.info(f"🔗 Initializing SharePoint connection...")
            graph_client = GraphAPIClient(
                client_id=self.graph_credentials["client_id"],
                client_secret=self.graph_credentials["client_secret"],
                tenant_id=self.graph_credentials["tenant_id"],
                sharepoint_config=self.sharepoint_config,
            )
            
            # Get item ID from URL if needed
            if sharepoint_url:
                logger.info(f"📋 Resolving SharePoint URL to item ID...")
                item_id = graph_client.get_sharepoint_item_id_from_url(sharepoint_url)
                logger.debug(f"Resolved item ID: {item_id}")
            else:
                item_id = sharepoint_item_id
                logger.info(f"📋 Using provided item ID: {item_id}")
            
            # Download file to memory
            logger.info(f"⬇️ Downloading file from SharePoint...")
            file_data = graph_client.download_file_to_memory(item_id)
            
            # Create BytesIO object
            excel_file = io.BytesIO(file_data)
            excel_file.filename = default_filename
            
            logger.info(
                f"✅ Successfully downloaded SharePoint file: {len(file_data):,} bytes"
            )
            
            return excel_file
            
        except Exception as e:
            error_msg = f"Failed to download file from SharePoint: {e}"
            logger.error(f"❌ {error_msg}")
            raise ValidationError(error_msg)
    
    def get_credentials_summary(self) -> Dict[str, Any]:
        """Get summary of current credentials status for debugging.
        
        Returns:
            Dictionary with credential status information
        """
        return {
            "sharepoint_enabled": self.sharepoint_config.get("enabled", False),
            "tenant_id_configured": bool(self.sharepoint_config.get("tenant_id")),
            "graph_credentials_loaded": bool(self.graph_credentials),
            "client_id_present": bool(self.graph_credentials and self.graph_credentials.get("client_id")),
            "client_secret_present": bool(self.graph_credentials and self.graph_credentials.get("client_secret")),
            "tenant_id_present": bool(self.graph_credentials and self.graph_credentials.get("tenant_id")),
        }