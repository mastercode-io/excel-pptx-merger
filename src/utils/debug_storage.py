"""Debug storage service for persistent debugging data in Google Cloud Storage."""

import os
import json
import datetime
import logging
from typing import Dict, Any, Optional, List
from .storage import StorageFactory

logger = logging.getLogger(__name__)


class DebugStorageService:
    """Service for storing debug data persistently in Google Cloud Storage."""

    def __init__(self):
        """Initialize the debug storage service."""
        self.bucket_name = os.getenv("DEBUG_STORAGE_BUCKET", "excel-pptx-merger-storage")
        self.debug_prefix = "debug"
        self.enabled = os.getenv("ENABLE_DEBUG_SAVING", "true").lower() == "true"

        # Initialize storage backend
        try:
            self.storage = StorageFactory.get_storage_backend()
            logger.info(f"📦 Debug storage initialized: bucket={self.bucket_name}")
            print(f"📦 DEBUG STORAGE: Initialized with bucket={self.bucket_name}")
            print(f"📦 DEBUG STORAGE: Backend type={type(self.storage).__name__}")

            # Test storage connectivity
            self._test_storage_connectivity()

        except Exception as e:
            logger.error(f"❌ Failed to initialize debug storage: {e}")
            print(f"❌ DEBUG STORAGE INIT FAILED: {e}")
            self.storage = None
            self.enabled = False

    def save_debug_data(
        self,
        request_data: Dict[str, Any],
        extracted_data: Dict[str, Any],
        session_id: str,
        context: str = "request"
    ) -> Optional[Dict[str, str]]:
        """Save complete debug data to persistent storage.

        Args:
            request_data: Original request information (excluding file content)
            extracted_data: Extracted data from Excel processing
            session_id: Unique session identifier
            context: Context of the debug save

        Returns:
            Dictionary with debug file information or None if saving failed
        """
        if not self.enabled or not self.storage:
            logger.debug("Debug storage disabled or not available")
            return None

        try:
            # Generate filename with timestamp and session ID
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"debug_{timestamp}_{session_id}_{context}.json"

            # Create comprehensive debug data
            debug_data = {
                "metadata": {
                    "timestamp": datetime.datetime.now().isoformat(),
                    "session_id": session_id,
                    "context": context,
                    "version": "1.0"
                },
                "request_info": request_data,
                "extraction_results": {
                    "extracted_data": extracted_data,
                    "formula_issues_detected": self._detect_formulas(extracted_data),
                    "formula_samples": self._extract_formula_samples(extracted_data),
                    "data_size": len(json.dumps(extracted_data, default=str))
                },
                "environment": {
                    "enable_debug_saving": self.enabled,
                    "force_formula_calculation": os.getenv("FORCE_FORMULA_CALCULATION", "false"),
                    "clean_excel_quotes": os.getenv("CLEAN_EXCEL_QUOTES", "true"),
                    "development_mode": os.getenv("DEVELOPMENT_MODE", "false"),
                    "storage_backend": os.getenv("STORAGE_BACKEND", "LOCAL")
                }
            }

            # Save to GCS
            debug_path = f"{self.debug_prefix}/{filename}"
            debug_content = json.dumps(debug_data, indent=2, default=str)

            # Use storage backend to save file
            saved_path = self.storage.save_file(debug_path, debug_content.encode('utf-8'))

            logger.info(f"📁 Debug data saved: {saved_path}")

            # Generate public URL if possible
            public_url = None
            try:
                public_url = self.storage.get_public_url(debug_path, expiration_seconds=86400)  # 24 hours
            except Exception as e:
                logger.debug(f"Could not generate public URL: {e}")

            return {
                "filename": filename,
                "path": saved_path,
                "bucket": self.bucket_name,
                "public_url": public_url,
                "timestamp": timestamp,
                "context": context,
                "size_bytes": len(debug_content)
            }

        except Exception as e:
            logger.error(f"Failed to save debug data: {e}")
            print(f"❌ DEBUG SAVE ERROR: {e}")
            print(f"❌ Storage backend: {type(self.storage).__name__ if self.storage else 'None'}")
            print(f"❌ Storage enabled: {self.enabled}")
            print(f"❌ Bucket: {self.bucket_name}")
            return None

    def list_debug_files(self, limit: int = 50) -> List[Dict[str, Any]]:
        """List available debug files.

        Args:
            limit: Maximum number of files to return

        Returns:
            List of debug file information
        """
        if not self.enabled or not self.storage:
            return []

        try:
            # List files in debug directory
            debug_files = self.storage.list_directory(self.debug_prefix)

            # Sort by date (newest first) and limit
            debug_files.sort(reverse=True)
            debug_files = debug_files[:limit]

            # Get file information
            file_info = []
            for filename in debug_files:
                if filename.endswith('.json'):
                    filepath = f"{self.debug_prefix}/{filename}"
                    try:
                        # Try to get public URL
                        public_url = self.storage.get_public_url(filepath, expiration_seconds=3600)
                    except:
                        public_url = None

                    file_info.append({
                        "filename": filename,
                        "path": filepath,
                        "public_url": public_url,
                        "exists": self.storage.file_exists(filepath)
                    })

            return file_info

        except Exception as e:
            logger.error(f"Failed to list debug files: {e}")
            return []

    def get_debug_file(self, filename: str) -> Optional[bytes]:
        """Retrieve a debug file by filename.

        Args:
            filename: Name of the debug file

        Returns:
            File content as bytes or None if not found
        """
        if not self.enabled or not self.storage:
            return None

        try:
            filepath = f"{self.debug_prefix}/{filename}"

            if not self.storage.file_exists(filepath):
                logger.warning(f"Debug file not found: {filename}")
                return None

            return self.storage.read_file(filepath)

        except Exception as e:
            logger.error(f"Failed to retrieve debug file {filename}: {e}")
            return None

    def _detect_formulas(self, data: Dict[str, Any]) -> bool:
        """Detect if data contains Excel formulas."""
        try:
            from .validation import detect_formula_extraction_issue
            return detect_formula_extraction_issue(data)
        except:
            return False

    def _extract_formula_samples(self, data: Dict[str, Any], max_samples: int = 5) -> List[str]:
        """Extract sample formulas from the data."""
        samples = []

        def find_formulas(obj):
            if len(samples) >= max_samples:
                return

            if isinstance(obj, dict):
                for value in obj.values():
                    find_formulas(value)
            elif isinstance(obj, list):
                for item in obj:
                    find_formulas(item)
            elif isinstance(obj, str) and obj.startswith("="):
                samples.append(obj)

        find_formulas(data)
        return samples

    def cleanup_old_files(self, days: int = 30) -> int:
        """Clean up debug files older than specified days.

        Args:
            days: Number of days to keep files

        Returns:
            Number of files deleted
        """
        if not self.enabled or not self.storage:
            return 0

        try:
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=days)
            deleted_count = 0

            debug_files = self.storage.list_directory(self.debug_prefix)

            for filename in debug_files:
                if filename.endswith('.json'):
                    # Extract date from filename (debug_YYYYMMDD_HHMMSS_...)
                    try:
                        date_part = filename.split('_')[1]  # YYYYMMDD
                        file_date = datetime.datetime.strptime(date_part, "%Y%m%d")

                        if file_date < cutoff_date:
                            filepath = f"{self.debug_prefix}/{filename}"
                            if self.storage.delete_file(filepath):
                                deleted_count += 1
                                logger.debug(f"Deleted old debug file: {filename}")
                    except (IndexError, ValueError):
                        # Skip files with invalid date format
                        continue

            if deleted_count > 0:
                logger.info(f"Cleaned up {deleted_count} old debug files")

            return deleted_count

        except Exception as e:
            logger.error(f"Failed to cleanup old debug files: {e}")
            return 0

    def _test_storage_connectivity(self) -> None:
        """Test storage connectivity on initialization."""
        try:
            if not self.storage:
                print("⚠️ DEBUG STORAGE: No storage backend available")
                return

            # Test basic operations
            test_path = f"{self.debug_prefix}/test_connectivity.json"
            test_data = b'{"test": "connectivity"}'

            # Try to save a test file
            saved_path = self.storage.save_file(test_path, test_data)
            print(f"✅ DEBUG STORAGE: Test file saved to {saved_path}")

            # Try to read it back
            if self.storage.file_exists(test_path):
                read_data = self.storage.read_file(test_path)
                if read_data == test_data:
                    print("✅ DEBUG STORAGE: Read test successful")
                else:
                    print("⚠️ DEBUG STORAGE: Read test failed - data mismatch")
            else:
                print("⚠️ DEBUG STORAGE: Test file not found after save")

            # Clean up test file
            if self.storage.delete_file(test_path):
                print("✅ DEBUG STORAGE: Test cleanup successful")
            else:
                print("⚠️ DEBUG STORAGE: Test cleanup failed")

            print("✅ DEBUG STORAGE: Connectivity test completed successfully")

        except Exception as e:
            logger.error(f"Storage connectivity test failed: {e}")
            print(f"❌ DEBUG STORAGE CONNECTIVITY FAILED: {e}")
            print("❌ Debug storage may not work properly")


# Global instance
debug_storage = DebugStorageService()