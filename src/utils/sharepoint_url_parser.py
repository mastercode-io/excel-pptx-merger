"""SharePoint URL parsing utilities for auto-resolution of site and drive information."""

import re
import logging
from typing import Dict, Optional, Tuple
from urllib.parse import urlparse, unquote

logger = logging.getLogger(__name__)


class SharePointUrlParser:
    """Utility class for parsing SharePoint URLs and extracting site/drive information."""

    # Common SharePoint URL patterns
    SHAREPOINT_URL_PATTERN = re.compile(
        r"https://([^/]+)\.sharepoint\.com/sites/([^/]+)/([^/]+)/(.+)$"
    )

    # Doc.aspx pattern for direct document access
    DOC_ASPX_PATTERN = re.compile(
        r"https://([^/]+)\.sharepoint\.com/:([xpwb]):/r/_layouts/15/Doc\.aspx\?sourcedoc=([^&]+)"
    )

    # Document library mappings
    LIBRARY_MAPPINGS = {
        "Shared Documents": "Documents",
        "Shared%20Documents": "Documents",
        "Documents": "Documents",
        "Site Assets": "SiteAssets",
        "Style Library": "Style%20Library",
        "Form Templates": "FormServerTemplates",
        "Site Pages": "SitePages",
    }

    def __init__(self):
        """Initialize SharePoint URL parser."""
        pass

    def parse_sharepoint_url(self, url: str) -> Optional[Dict[str, str]]:
        """Parse SharePoint URL and extract components.

        Args:
            url: SharePoint URL to parse

        Returns:
            Dictionary with parsed components or None if not a valid SharePoint URL
        """
        try:
            parsed = urlparse(url)

            # Basic validation
            if not parsed.netloc.endswith(".sharepoint.com"):
                logger.warning(f"URL does not appear to be a SharePoint URL: {url}")
                return None

            # Extract tenant from hostname
            tenant_match = re.match(r"([^.]+)\.sharepoint\.com", parsed.netloc)
            if not tenant_match:
                logger.warning(f"Could not extract tenant from URL: {url}")
                return None

            tenant = tenant_match.group(1)

            # Check if it's a Doc.aspx URL
            doc_aspx_match = self.DOC_ASPX_PATTERN.match(url)
            if doc_aspx_match:
                return self._parse_doc_aspx_url(url, tenant, doc_aspx_match)

            # Otherwise, try standard path pattern
            path_parts = [part for part in parsed.path.split("/") if part]

            if len(path_parts) < 4 or path_parts[0] != "sites":
                logger.warning(
                    f"URL does not follow expected SharePoint sites pattern: {url}"
                )
                return None

            site_name = path_parts[1]
            library_name = unquote(path_parts[2])
            file_path = "/".join(path_parts[3:])

            # Normalize library name
            normalized_library = self._normalize_library_name(library_name)

            result = {
                "tenant": tenant,
                "site_name": site_name,
                "library_name": library_name,
                "normalized_library": normalized_library,
                "file_path": file_path,
                "full_url": url,
                "hostname": parsed.netloc,
                "url_type": "path",
            }

            logger.debug(f"Parsed SharePoint URL: {result}")
            return result

        except Exception as e:
            logger.error(f"Error parsing SharePoint URL '{url}': {e}")
            return None

    def _parse_doc_aspx_url(
        self, url: str, tenant: str, match: re.Match
    ) -> Dict[str, str]:
        """Parse Doc.aspx URL format.

        Args:
            url: Full URL
            tenant: Extracted tenant name
            match: Regex match object from DOC_ASPX_PATTERN

        Returns:
            Dictionary with parsed components
        """
        file_type_code = match.group(2)
        sourcedoc = match.group(3)

        # Map file type codes to types
        file_type_map = {"x": "excel", "p": "powerpoint", "w": "word", "b": "onenote"}

        file_type = file_type_map.get(file_type_code, "unknown")

        # Extract filename from query parameters if available
        parsed = urlparse(url)
        params = dict(
            param.split("=") for param in parsed.query.split("&") if "=" in param
        )
        filename = unquote(params.get("file", ""))

        # Clean up sourcedoc (remove URL encoding)
        sourcedoc_clean = unquote(sourcedoc).strip("{}")

        result = {
            "tenant": tenant,
            "hostname": f"{tenant}.sharepoint.com",
            "full_url": url,
            "url_type": "doc_aspx",
            "file_type": file_type,
            "sourcedoc": sourcedoc_clean,
            "filename": filename,
            "requires_shares_api": True,
        }

        logger.debug(f"Parsed Doc.aspx URL: {result}")
        return result

    def _normalize_library_name(self, library_name: str) -> str:
        """Normalize library name to standard format."""
        # Handle URL encoding
        decoded = unquote(library_name)

        # Check for known mappings
        if decoded in self.LIBRARY_MAPPINGS:
            return self.LIBRARY_MAPPINGS[decoded]

        # Default: return as-is
        return decoded

    def extract_tenant_id_from_url(self, url: str) -> Optional[str]:
        """Extract tenant identifier from SharePoint URL.

        Args:
            url: SharePoint URL

        Returns:
            Tenant identifier or None if not found
        """
        parsed = self.parse_sharepoint_url(url)
        if parsed:
            return parsed["tenant"]
        return None

    def validate_sharepoint_url(self, url: str) -> bool:
        """Validate if URL is a valid SharePoint URL.

        Args:
            url: URL to validate

        Returns:
            True if valid SharePoint URL, False otherwise
        """
        return self.parse_sharepoint_url(url) is not None

    def extract_site_and_drive_info(
        self, url: str
    ) -> Tuple[Optional[str], Optional[str]]:
        """Extract site name and library name from SharePoint URL.

        Args:
            url: SharePoint URL

        Returns:
            Tuple of (site_name, library_name) or (None, None) if parsing fails
        """
        parsed = self.parse_sharepoint_url(url)
        if parsed:
            return parsed["site_name"], parsed["normalized_library"]
        return None, None


def parse_sharepoint_url(url: str) -> Optional[Dict[str, str]]:
    """Convenience function for parsing SharePoint URLs."""
    parser = SharePointUrlParser()
    return parser.parse_sharepoint_url(url)


def extract_tenant_from_url(url: str) -> Optional[str]:
    """Convenience function for extracting tenant from SharePoint URL."""
    parser = SharePointUrlParser()
    return parser.extract_tenant_id_from_url(url)


def validate_sharepoint_url(url: str) -> bool:
    """Convenience function for validating SharePoint URLs."""
    parser = SharePointUrlParser()
    return parser.validate_sharepoint_url(url)
