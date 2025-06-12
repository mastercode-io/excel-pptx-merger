"""Storage abstraction layer for handling files in different storage backends.

This module provides a unified interface for file operations across different storage backends:
- Local filesystem (for development)
- Google Cloud Storage (for production)

The storage backend is determined by environment variables.
"""

import os
import logging
import tempfile
from abc import ABC, abstractmethod
from typing import BinaryIO, Optional, Union, List, Dict, Any
from pathlib import Path
import shutil

# Import Google Cloud Storage only if available
try:
    from google.cloud import storage
    from google.cloud.exceptions import NotFound
    GCS_AVAILABLE = True
except ImportError:
    GCS_AVAILABLE = False

logger = logging.getLogger(__name__)


class StorageBackend(ABC):
    """Abstract base class for storage backends."""

    @abstractmethod
    def save_file(self, file_path: str, content: Union[bytes, BinaryIO]) -> str:
        """Save file content to the specified path."""
        pass

    @abstractmethod
    def read_file(self, file_path: str) -> bytes:
        """Read file content from the specified path."""
        pass

    @abstractmethod
    def delete_file(self, file_path: str) -> bool:
        """Delete a file at the specified path."""
        pass

    @abstractmethod
    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists at the specified path."""
        pass

    @abstractmethod
    def create_directory(self, directory_path: str) -> str:
        """Create a directory at the specified path."""
        pass

    @abstractmethod
    def delete_directory(self, directory_path: str) -> bool:
        """Delete a directory and all its contents."""
        pass

    @abstractmethod
    def list_directory(self, directory_path: str) -> List[str]:
        """List all files in a directory."""
        pass

    @abstractmethod
    def get_public_url(self, file_path: str, expiration_seconds: int = 3600) -> Optional[str]:
        """Get a public URL for accessing the file."""
        pass


class LocalStorageBackend(StorageBackend):
    """Local filesystem storage backend for development."""

    def __init__(self, base_directory: Optional[str] = None):
        """Initialize local storage backend.
        
        Args:
            base_directory: Base directory for all operations. If None, uses system temp directory.
        """
        if base_directory:
            self.base_directory = base_directory
            os.makedirs(self.base_directory, exist_ok=True)
        else:
            self.base_directory = tempfile.gettempdir()
        
        logger.info(f"LocalStorageBackend initialized with base directory: {self.base_directory}")

    def _get_full_path(self, path: str) -> str:
        """Get the full path by joining with the base directory."""
        # If path is already absolute, return it as is
        if os.path.isabs(path):
            return path
        return os.path.join(self.base_directory, path)

    def save_file(self, file_path: str, content: Union[bytes, BinaryIO]) -> str:
        """Save file content to the specified path."""
        full_path = self._get_full_path(file_path)
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        
        if isinstance(content, bytes):
            with open(full_path, 'wb') as f:
                f.write(content)
        else:
            # Assume it's a file-like object
            content.seek(0)
            with open(full_path, 'wb') as f:
                shutil.copyfileobj(content, f)
        
        logger.debug(f"Saved file to local storage: {full_path}")
        return full_path

    def read_file(self, file_path: str) -> bytes:
        """Read file content from the specified path."""
        full_path = self._get_full_path(file_path)
        
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"File not found: {full_path}")
        
        with open(full_path, 'rb') as f:
            content = f.read()
        
        logger.debug(f"Read file from local storage: {full_path}")
        return content

    def delete_file(self, file_path: str) -> bool:
        """Delete a file at the specified path."""
        full_path = self._get_full_path(file_path)
        
        if not os.path.exists(full_path):
            logger.warning(f"File not found for deletion: {full_path}")
            return False
        
        os.remove(full_path)
        logger.debug(f"Deleted file from local storage: {full_path}")
        return True

    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists at the specified path."""
        full_path = self._get_full_path(file_path)
        return os.path.isfile(full_path)

    def create_directory(self, directory_path: str) -> str:
        """Create a directory at the specified path."""
        full_path = self._get_full_path(directory_path)
        os.makedirs(full_path, exist_ok=True)
        logger.debug(f"Created directory in local storage: {full_path}")
        return full_path

    def delete_directory(self, directory_path: str) -> bool:
        """Delete a directory and all its contents."""
        full_path = self._get_full_path(directory_path)
        
        if not os.path.exists(full_path):
            logger.warning(f"Directory not found for deletion: {full_path}")
            return False
        
        shutil.rmtree(full_path)
        logger.debug(f"Deleted directory from local storage: {full_path}")
        return True

    def list_directory(self, directory_path: str) -> List[str]:
        """List all files in a directory."""
        full_path = self._get_full_path(directory_path)
        
        if not os.path.exists(full_path):
            logger.warning(f"Directory not found for listing: {full_path}")
            return []
        
        return [os.path.join(directory_path, f) for f in os.listdir(full_path)]

    def get_public_url(self, file_path: str, expiration_seconds: int = 3600) -> Optional[str]:
        """Get a public URL for accessing the file.
        
        For local storage, this returns a file:// URL which is only useful for local access.
        """
        full_path = self._get_full_path(file_path)
        
        if not os.path.exists(full_path):
            return None
        
        # For local files, just return a file:// URL
        return f"file://{os.path.abspath(full_path)}"


class GCSStorageBackend(StorageBackend):
    """Google Cloud Storage backend for production."""

    def __init__(self, bucket_name: str, base_prefix: str = ""):
        """Initialize GCS storage backend.
        
        Args:
            bucket_name: Name of the GCS bucket
            base_prefix: Base prefix for all objects in the bucket
        """
        if not GCS_AVAILABLE:
            raise ImportError("Google Cloud Storage library not available. "
                             "Install with 'pip install google-cloud-storage'")
        
        self.bucket_name = bucket_name
        self.base_prefix = base_prefix.rstrip('/') + '/' if base_prefix else ""
        
        # Initialize GCS client
        self.client = storage.Client()
        self.bucket = self.client.bucket(bucket_name)
        
        logger.info(f"GCSStorageBackend initialized with bucket: {bucket_name}, "
                   f"base prefix: {self.base_prefix}")

    def _get_full_path(self, path: str) -> str:
        """Get the full object path by joining with the base prefix."""
        # Remove leading slash if present
        path = path.lstrip('/')
        return f"{self.base_prefix}{path}"

    def save_file(self, file_path: str, content: Union[bytes, BinaryIO]) -> str:
        """Save file content to the specified path in GCS."""
        object_path = self._get_full_path(file_path)
        blob = self.bucket.blob(object_path)
        
        if isinstance(content, bytes):
            blob.upload_from_string(content, content_type=self._guess_content_type(file_path))
        else:
            # Assume it's a file-like object
            content.seek(0)
            blob.upload_from_file(content, content_type=self._guess_content_type(file_path))
        
        logger.debug(f"Saved file to GCS: gs://{self.bucket_name}/{object_path}")
        return f"gs://{self.bucket_name}/{object_path}"

    def read_file(self, file_path: str) -> bytes:
        """Read file content from the specified path in GCS."""
        object_path = self._get_full_path(file_path)
        blob = self.bucket.blob(object_path)
        
        try:
            content = blob.download_as_bytes()
            logger.debug(f"Read file from GCS: gs://{self.bucket_name}/{object_path}")
            return content
        except NotFound:
            raise FileNotFoundError(f"File not found in GCS: gs://{self.bucket_name}/{object_path}")

    def delete_file(self, file_path: str) -> bool:
        """Delete a file at the specified path in GCS."""
        object_path = self._get_full_path(file_path)
        blob = self.bucket.blob(object_path)
        
        try:
            blob.delete()
            logger.debug(f"Deleted file from GCS: gs://{self.bucket_name}/{object_path}")
            return True
        except NotFound:
            logger.warning(f"File not found for deletion in GCS: gs://{self.bucket_name}/{object_path}")
            return False

    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists at the specified path in GCS."""
        object_path = self._get_full_path(file_path)
        blob = self.bucket.blob(object_path)
        return blob.exists()

    def create_directory(self, directory_path: str) -> str:
        """Create a directory at the specified path in GCS.
        
        Note: GCS doesn't have actual directories, but we can create a placeholder object
        with a trailing slash to simulate a directory.
        """
        # Ensure path ends with slash
        directory_path = directory_path.rstrip('/') + '/'
        object_path = self._get_full_path(directory_path)
        
        # Create an empty object with the directory path
        blob = self.bucket.blob(object_path)
        blob.upload_from_string('')
        
        logger.debug(f"Created directory in GCS: gs://{self.bucket_name}/{object_path}")
        return f"gs://{self.bucket_name}/{object_path}"

    def delete_directory(self, directory_path: str) -> bool:
        """Delete a directory and all its contents in GCS."""
        # Ensure path ends with slash
        directory_path = directory_path.rstrip('/') + '/'
        prefix = self._get_full_path(directory_path)
        
        blobs = list(self.bucket.list_blobs(prefix=prefix))
        if not blobs:
            logger.warning(f"Directory not found or empty in GCS: gs://{self.bucket_name}/{prefix}")
            return False
        
        # Delete all objects with the prefix
        for blob in blobs:
            blob.delete()
        
        logger.debug(f"Deleted directory from GCS: gs://{self.bucket_name}/{prefix}")
        return True

    def list_directory(self, directory_path: str) -> List[str]:
        """List all files in a directory in GCS."""
        # Ensure path ends with slash
        directory_path = directory_path.rstrip('/') + '/'
        prefix = self._get_full_path(directory_path)
        
        blobs = list(self.bucket.list_blobs(prefix=prefix))
        
        # Convert GCS paths to relative paths
        result = []
        prefix_len = len(prefix)
        for blob in blobs:
            if blob.name != prefix:  # Skip the directory placeholder
                result.append(directory_path + blob.name[prefix_len:])
        
        return result

    def get_public_url(self, file_path: str, expiration_seconds: int = 3600) -> Optional[str]:
        """Get a signed URL for accessing the file in GCS."""
        object_path = self._get_full_path(file_path)
        blob = self.bucket.blob(object_path)
        
        if not blob.exists():
            return None
        
        # Generate a signed URL with expiration
        url = blob.generate_signed_url(
            version="v4",
            expiration=expiration_seconds,
            method="GET"
        )
        
        return url

    def _guess_content_type(self, file_path: str) -> str:
        """Guess the content type based on file extension."""
        import mimetypes
        content_type, _ = mimetypes.guess_type(file_path)
        return content_type or 'application/octet-stream'


class StorageFactory:
    """Factory for creating storage backend instances based on environment settings."""
    
    @staticmethod
    def get_storage_backend() -> StorageBackend:
        """Get the appropriate storage backend based on environment settings."""
        storage_type = os.getenv('STORAGE_BACKEND', 'LOCAL').upper()
        
        if storage_type == 'GCS':
            bucket_name = os.getenv('GCS_BUCKET_NAME')
            if not bucket_name:
                logger.warning("GCS_BUCKET_NAME not set, falling back to local storage")
                return LocalStorageBackend()
            
            base_prefix = os.getenv('GCS_BASE_PREFIX', '')
            
            try:
                return GCSStorageBackend(bucket_name, base_prefix)
            except ImportError:
                logger.warning("Google Cloud Storage library not available, falling back to local storage")
                return LocalStorageBackend()
        else:
            # Default to local storage
            base_directory = os.getenv('LOCAL_STORAGE_DIR')
            return LocalStorageBackend(base_directory)
