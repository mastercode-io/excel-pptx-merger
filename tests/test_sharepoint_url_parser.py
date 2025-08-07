"""Tests for SharePoint URL parser."""

import pytest
from src.utils.sharepoint_url_parser import SharePointUrlParser


class TestSharePointUrlParser:
    """Test SharePoint URL parsing functionality."""

    def setup_method(self):
        """Setup test fixtures."""
        self.parser = SharePointUrlParser()

    def test_root_doc_aspx_url_excel(self):
        """Test parsing of root-level Doc.aspx URLs without /:x:/r/ prefix for Excel files."""
        # Test case matching the user's URL
        url = "https://thetrademarkhelpline.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B75EF7D72-101B-4391-91B2-5461A582E6AA%7D&file=28-07-25%20-%20Sistren%20Shirley%20-%20Sistren%20Shirley.xlsx&action=default&mobileredirect=true"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["tenant"] == "thetrademarkhelpline"
        assert result["url_type"] == "root_doc_aspx"
        assert result["file_type"] == "excel"
        assert result["filename"] == "28-07-25 - Sistren Shirley - Sistren Shirley.xlsx"
        assert result["requires_shares_api"] == True
        assert result["sourcedoc"] == "75EF7D72-101B-4391-91B2-5461A582E6AA"
        assert result["hostname"] == "thetrademarkhelpline.sharepoint.com"

    def test_root_doc_aspx_url_powerpoint(self):
        """Test parsing of root-level Doc.aspx URLs for PowerPoint files."""
        url = "https://example.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BABCDEF12-3456-7890-ABCD-EF1234567890%7D&file=Presentation.pptx"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["tenant"] == "example"
        assert result["url_type"] == "root_doc_aspx"
        assert result["file_type"] == "powerpoint"
        assert result["filename"] == "Presentation.pptx"
        assert result["requires_shares_api"] == True
        assert result["sourcedoc"] == "ABCDEF12-3456-7890-ABCD-EF1234567890"

    def test_root_doc_aspx_url_word(self):
        """Test parsing of root-level Doc.aspx URLs for Word files."""
        url = "https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B12345678-90AB-CDEF-1234-567890ABCDEF%7D&file=Document%20Name.docx"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["tenant"] == "contoso"
        assert result["url_type"] == "root_doc_aspx"
        assert result["file_type"] == "word"
        assert result["filename"] == "Document Name.docx"  # URL decoded
        assert result["requires_shares_api"] == True

    def test_root_doc_aspx_url_without_file_param(self):
        """Test parsing of root-level Doc.aspx URLs without file parameter."""
        url = "https://mycompany.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7BAAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE%7D"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["tenant"] == "mycompany"
        assert result["url_type"] == "root_doc_aspx"
        assert result["file_type"] == "unknown"  # No file extension to determine type
        assert result["filename"] == ""
        assert result["requires_shares_api"] == True

    def test_traditional_doc_aspx_url_still_works(self):
        """Test that traditional Doc.aspx URLs with /:x:/r/ prefix still work."""
        url = "https://example.sharepoint.com/:x:/r/_layouts/15/Doc.aspx?sourcedoc=%7B12345678-90AB-CDEF-1234-567890ABCDEF%7D"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["url_type"] == "doc_aspx"  # Different from root_doc_aspx
        assert result["file_type"] == "excel"  # Determined by :x: prefix
        assert result["requires_shares_api"] == True

    def test_standard_sites_url_still_works(self):
        """Test that standard /sites/ URLs still work."""
        url = "https://example.sharepoint.com/sites/TeamSite/Shared%20Documents/Report.xlsx"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["tenant"] == "example"
        assert result["site_name"] == "TeamSite"
        assert result["library_name"] == "Shared Documents"
        assert result["file_path"] == "Report.xlsx"
        assert result["url_type"] == "path"

    def test_invalid_sharepoint_url(self):
        """Test that invalid URLs return None."""
        invalid_urls = [
            "https://example.com/file.xlsx",  # Not SharePoint
            "http://example.sharepoint.com/file",  # HTTP not HTTPS
            "https://sharepoint.com/file",  # No tenant
            "not_a_url",  # Not a URL at all
        ]
        
        for url in invalid_urls:
            result = self.parser.parse_sharepoint_url(url)
            assert result is None

    def test_url_with_special_characters(self):
        """Test parsing URLs with special characters in filename."""
        url = "https://test.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B11111111-2222-3333-4444-555555555555%7D&file=File%20%26%20Name%20%23%201.xlsx"
        
        result = self.parser.parse_sharepoint_url(url)
        
        assert result is not None
        assert result["filename"] == "File & Name # 1.xlsx"  # Properly decoded
        assert result["file_type"] == "excel"

    def test_extract_tenant_from_url(self):
        """Test extracting tenant from various URL formats."""
        test_cases = [
            ("https://contoso.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B123%7D", "contoso"),
            ("https://mycompany.sharepoint.com/sites/team/Documents/file.xlsx", "mycompany"),
            ("https://example.sharepoint.com/:x:/r/_layouts/15/Doc.aspx?sourcedoc=%7B456%7D", "example"),
        ]
        
        for url, expected_tenant in test_cases:
            tenant = self.parser.extract_tenant_id_from_url(url)
            assert tenant == expected_tenant

    def test_validate_sharepoint_url(self):
        """Test URL validation."""
        valid_urls = [
            "https://test.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7B123%7D",
            "https://test.sharepoint.com/sites/team/Documents/file.xlsx",
            "https://test.sharepoint.com/:x:/r/_layouts/15/Doc.aspx?sourcedoc=%7B456%7D",
        ]
        
        for url in valid_urls:
            assert self.parser.validate_sharepoint_url(url) == True
        
        invalid_urls = [
            "https://example.com/file.xlsx",
            "not_a_url",
        ]
        
        for url in invalid_urls:
            assert self.parser.validate_sharepoint_url(url) == False


if __name__ == "__main__":
    pytest.main([__file__, "-v"])