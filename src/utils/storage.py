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
from datetime import timedelta

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

    # Standard folder structure
    INPUT_FOLDER = "input"
    OUTPUT_FOLDER = "output"
    IMAGES_FOLDER = "images"
    DEBUG_FOLDER = "debug"

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
    def get_public_url(
        self, file_path: str, expiration_seconds: int = 3600
    ) -> Optional[str]:
        """Get a public URL for accessing the file."""
        pass

    @abstractmethod
    def initialize_folder_structure(self, base_path: str) -> None:
        """Initialize the standard folder structure."""
        pass

    def get_input_path(self, base_path: str, filename: str) -> str:
        """Get path for input files."""
        return self._join_paths(base_path, self.INPUT_FOLDER, filename)

    def get_output_path(self, base_path: str, filename: str) -> str:
        """Get path for output files."""
        return self._join_paths(base_path, self.OUTPUT_FOLDER, filename)

    def get_image_path(self, base_path: str, filename: str) -> str:
        """Get path for image files."""
        return self._join_paths(base_path, self.IMAGES_FOLDER, filename)

    def get_debug_path(self, base_path: str, filename: str) -> str:
        """Get path for debug files."""
        return self._join_paths(base_path, self.DEBUG_FOLDER, filename)

    @abstractmethod
    def _join_paths(self, *paths: str) -> str:
        """Join path components in a backend-specific way."""
        pass


class LocalStorageBackend(StorageBackend):
    """Local filesystem storage backend for development."""

    def __init__(self, base_directory: Optional[str] = None):
        """Initialize local storage backend.

        Args:
            base_directory: Base directory for all operations. If None, uses system temp directory.
        """
        if base_directory:
            # Convert relative paths to absolute paths
            if not os.path.isabs(base_directory):
                # Get the absolute path relative to the current working directory
                self.base_directory = os.path.abspath(base_directory)
            else:
                self.base_directory = base_directory
            os.makedirs(self.base_directory, exist_ok=True)
        else:
            self.base_directory = tempfile.gettempdir()

        logger.info(
            f"LocalStorageBackend initialized with base directory: {self.base_directory}"
        )

    def _get_full_path(self, path: str) -> str:
        """Get the full path by joining with the base directory.

        Args:
            path: Relative or absolute path

        Returns:
            Absolute path
        """
        # If path is already absolute, return it as is
        if os.path.isabs(path):
            return path
        return os.path.join(self.base_directory, path)

    def save_file(self, file_path: str, content: Union[bytes, BinaryIO]) -> str:
        """Save file content to the specified path."""
        full_path = self._get_full_path(file_path)

        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(full_path), exist_ok=True)

        # Handle both bytes and file-like objects
        if isinstance(content, bytes):
            with open(full_path, "wb") as f:
                f.write(content)
        else:
            # Assume it's a file-like object
            content.seek(0)
            with open(full_path, "wb") as f:
                shutil.copyfileobj(content, f)

        return file_path

    def read_file(self, file_path: str) -> bytes:
        """Read file content from the specified path."""
        full_path = self._get_full_path(file_path)
        with open(full_path, "rb") as f:
            return f.read()

    def delete_file(self, file_path: str) -> bool:
        """Delete a file at the specified path."""
        full_path = self._get_full_path(file_path)
        if os.path.exists(full_path):
            os.remove(full_path)
            return True
        return False

    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists at the specified path."""
        full_path = self._get_full_path(file_path)
        return os.path.isfile(full_path)

    def create_directory(self, directory_path: str) -> str:
        """Create a directory at the specified path.

        Args:
            directory_path: Path to create

        Returns:
            Absolute path to the created directory
        """
        full_path = self._get_full_path(directory_path)
        os.makedirs(full_path, exist_ok=True)

        # Initialize the standard folder structure
        self.initialize_folder_structure(directory_path)

        return full_path

    def initialize_folder_structure(self, base_path: str) -> None:
        """Initialize standard folder structure for a session directory.

        Creates input, output, images, and debug subdirectories.

        Args:
            base_path: Base path for the session directory
        """
        # Get the absolute path
        full_base_path = self._get_full_path(base_path)

        # Create standard subdirectories
        for subdir in ["input", "output", "images", "debug"]:
            subdir_path = os.path.join(full_base_path, subdir)
            os.makedirs(subdir_path, exist_ok=True)
            logger.debug(f"Created directory: {subdir_path}")

    def get_input_path(self, base_path: str, filename: str) -> str:
        """Get the path for an input file.

        Args:
            base_path: Base path for the session directory
            filename: Name of the input file

        Returns:
            Absolute path to the input file
        """
        # Get the absolute base path
        full_base_path = self._get_full_path(base_path)

        # Create the input directory if it doesn't exist
        input_dir = os.path.join(full_base_path, "input")
        os.makedirs(input_dir, exist_ok=True)

        # Return the absolute path to the input file
        return os.path.join(input_dir, filename)

    def get_output_path(self, base_path: str, filename: str) -> str:
        """Get the path for an output file.

        Args:
            base_path: Base path for the session directory
            filename: Name of the output file

        Returns:
            Absolute path to the output file
        """
        # Get the absolute base path
        full_base_path = self._get_full_path(base_path)

        # Create the output directory if it doesn't exist
        output_dir = os.path.join(full_base_path, "output")
        os.makedirs(output_dir, exist_ok=True)

        # Return the absolute path to the output file
        return os.path.join(output_dir, filename)

    def get_image_path(self, base_path: str, filename: str) -> str:
        """Get the path for an image file.

        Args:
            base_path: Base path for the session directory
            filename: Name of the image file

        Returns:
            Absolute path to the image file
        """
        # Get the absolute base path
        full_base_path = self._get_full_path(base_path)

        # Create the images directory if it doesn't exist
        images_dir = os.path.join(full_base_path, "images")
        os.makedirs(images_dir, exist_ok=True)

        # Return the absolute path to the image file
        return os.path.join(images_dir, filename)

    def get_debug_path(self, base_path: str, filename: str) -> str:
        """Get the path for a debug file.

        Args:
            base_path: Base path for the session directory
            filename: Name of the debug file

        Returns:
            Absolute path to the debug file
        """
        # Get the absolute base path
        full_base_path = self._get_full_path(base_path)

        # Create the debug directory if it doesn't exist
        debug_dir = os.path.join(full_base_path, "debug")
        os.makedirs(debug_dir, exist_ok=True)

        # Return the absolute path to the debug file
        return os.path.join(debug_dir, filename)

    def delete_directory(self, directory_path: str) -> bool:
        """Delete a directory and all its contents."""
        full_path = self._get_full_path(directory_path)
        if os.path.exists(full_path):
            shutil.rmtree(full_path)
            return True
        return False

    def list_directory(self, directory_path: str) -> List[str]:
        """List all files in a directory."""
        full_path = self._get_full_path(directory_path)
        if not os.path.exists(full_path):
            return []

        result = []
        for root, _, files in os.walk(full_path):
            for file in files:
                file_path = os.path.join(root, file)
                # Convert to relative path
                rel_path = os.path.relpath(file_path, self.base_directory)
                result.append(rel_path)

        return result

    def get_public_url(
        self, file_path: str, expiration_seconds: int = 3600
    ) -> Optional[str]:
        """Get a public URL for accessing the file."""
        # Local storage doesn't support public URLs
        return None

    def _join_paths(self, *paths: str) -> str:
        """Join path components for local filesystem."""
        return os.path.join(*paths)


class GCSStorageBackend(StorageBackend):
    """Google Cloud Storage backend for production."""

    def __init__(self, bucket_name: str, base_prefix: str = ""):
        """Initialize GCS storage backend.

        Args:
            bucket_name: Name of the GCS bucket
            base_prefix: Base prefix for all objects in the bucket
        """
        if not GCS_AVAILABLE:
            raise ImportError(
                "Google Cloud Storage library not available. "
                "Install with 'pip install google-cloud-storage'"
            )

        self.bucket_name = bucket_name
        self.base_prefix = base_prefix.rstrip("/") + "/" if base_prefix else ""

        # Initialize GCS client
        self.client = storage.Client()
        self.bucket = self.client.bucket(bucket_name)

        logger.info(
            f"GCSStorageBackend initialized with bucket: {bucket_name}, "
            f"base prefix: {self.base_prefix}"
        )

    def _get_full_path(self, path: str) -> str:
        """Get the full object path by joining with the base prefix."""
        # Remove leading slash if present
        path = path.lstrip("/")
        return f"{self.base_prefix}{path}"

    def save_file(self, file_path: str, content: Union[bytes, BinaryIO]) -> str:
        """Save file content to the specified path in GCS."""
        full_path = self._get_full_path(file_path)
        blob = self.bucket.blob(full_path)

        # Set content type based on file extension
        content_type = self._guess_content_type(file_path)
        if content_type:
            blob.content_type = content_type

        # Handle both bytes and file-like objects
        if isinstance(content, bytes):
            blob.upload_from_string(content)
        else:
            # Assume it's a file-like object
            content.seek(0)
            blob.upload_from_file(content)

        return file_path

    def read_file(self, file_path: str) -> bytes:
        """Read file content from the specified path in GCS."""
        full_path = self._get_full_path(file_path)
        blob = self.bucket.blob(full_path)

        if not blob.exists():
            raise FileNotFoundError(f"File not found in GCS: {file_path}")

        return blob.download_as_bytes()

    def delete_file(self, file_path: str) -> bool:
        """Delete a file at the specified path in GCS."""
        full_path = self._get_full_path(file_path)
        blob = self.bucket.blob(full_path)

        if blob.exists():
            blob.delete()
            return True
        return False

    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists at the specified path in GCS."""
        full_path = self._get_full_path(file_path)
        blob = self.bucket.blob(full_path)
        return blob.exists()

    def create_directory(self, directory_path: str) -> str:
        """Create a directory at the specified path in GCS.

        Note: GCS doesn't have actual directories, but we can create a placeholder object
        with a trailing slash to simulate a directory.
        """
        # Ensure the path ends with a slash
        directory_path = directory_path.rstrip("/") + "/"
        full_path = self._get_full_path(directory_path)

        # Create a placeholder object
        blob = self.bucket.blob(full_path)
        blob.upload_from_string("", content_type="application/x-directory")

        # Initialize standard folder structure
        self.initialize_folder_structure(directory_path)

        return directory_path

    def delete_directory(self, directory_path: str) -> bool:
        """Delete a directory and all its contents in GCS."""
        # Ensure the path ends with a slash
        directory_path = directory_path.rstrip("/") + "/"
        full_path = self._get_full_path(directory_path)

        # List all blobs with the directory prefix
        blobs = list(self.bucket.list_blobs(prefix=full_path))

        if not blobs:
            return False

        # Delete all blobs
        for blob in blobs:
            blob.delete()

        return True

    def list_directory(self, directory_path: str) -> List[str]:
        """List all files in a directory in GCS."""
        # Ensure the path ends with a slash
        directory_path = directory_path.rstrip("/") + "/"
        full_path = self._get_full_path(directory_path)

        # List all blobs with the directory prefix
        blobs = list(self.bucket.list_blobs(prefix=full_path))

        # Extract relative paths
        result = []
        prefix_len = len(self.base_prefix)
        for blob in blobs:
            # Skip the directory placeholder itself
            if blob.name == full_path:
                continue

            # Remove base prefix to get relative path
            rel_path = blob.name[prefix_len:]
            result.append(rel_path)

        return result

    def get_public_url(
        self, file_path: str, expiration_seconds: int = 3600
    ) -> Optional[str]:
        """Get a signed URL for accessing the file in GCS."""
        full_path = self._get_full_path(file_path)
        blob = self.bucket.blob(full_path)

        if not blob.exists():
            return None

        # Generate a signed URL that expires after the specified time
        return blob.generate_signed_url(
            version="v4", expiration=timedelta(seconds=expiration_seconds), method="GET"
        )

    def initialize_folder_structure(self, base_path: str) -> None:
        """Initialize the standard folder structure for GCS."""
        # Ensure the base path ends with a slash
        base_path = base_path.rstrip("/") + "/"

        # Create placeholder objects for each folder
        for folder in [
            self.INPUT_FOLDER,
            self.OUTPUT_FOLDER,
            self.IMAGES_FOLDER,
            self.DEBUG_FOLDER,
        ]:
            folder_path = f"{base_path}{folder}/"
            full_path = self._get_full_path(folder_path)
            blob = self.bucket.blob(full_path)
            if not blob.exists():
                blob.upload_from_string("", content_type="application/x-directory")
                logger.debug(f"Created directory: {folder_path}")

    def _join_paths(self, *paths: str) -> str:
        """Join path components for GCS."""
        return "/".join(p.strip("/") for p in paths if p)

    def _guess_content_type(self, file_path: str) -> Optional[str]:
        """Guess the content type based on file extension."""
        ext = file_path.lower().split(".")[-1]
        content_types = {
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "json": "application/json",
            "txt": "text/plain",
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
        }
        return content_types.get(ext)


class StorageFactory:
    """Factory for creating storage backend instances based on environment settings."""

    @staticmethod
    def get_storage_backend() -> StorageBackend:
        """Get the appropriate storage backend based on environment settings."""
        storage_type = os.getenv("STORAGE_BACKEND", "LOCAL").upper()

        if storage_type == "GCS":
            bucket_name = os.getenv("GCS_BUCKET_NAME")
            if not bucket_name:
                logger.warning("GCS_BUCKET_NAME not set, falling back to local storage")
                return LocalStorageBackend()

            base_prefix = os.getenv("GCS_BASE_PREFIX", "")
            try:
                return GCSStorageBackend(bucket_name, base_prefix)
            except ImportError as e:
                logger.warning(f"Failed to initialize GCS backend: {e}")
                logger.warning("Falling back to local storage")
                return LocalStorageBackend()
        else:
            # Default to local storage
            local_dir = os.getenv("LOCAL_STORAGE_DIR")
            return LocalStorageBackend(local_dir)
