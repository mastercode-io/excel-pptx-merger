"""File handling utilities for Excel to PowerPoint Merger."""

import os
import shutil
import tempfile
from pathlib import Path
from typing import List, Optional, Union, BinaryIO
import logging
from werkzeug.datastructures import FileStorage

from .exceptions import FileProcessingError, ValidationError

logger = logging.getLogger(__name__)


def validate_file_extension(filename: str, allowed_extensions: List[str]) -> bool:
    """Validate file extension against allowed list."""
    if not filename:
        return False

    extension = Path(filename).suffix.lower().lstrip(".")
    return extension in [ext.lower().lstrip(".") for ext in allowed_extensions]


def validate_file_size(
    file_obj: Union[BinaryIO, FileStorage], max_size_mb: int
) -> bool:
    """Validate file size against maximum allowed size."""
    try:
        if hasattr(file_obj, "seek") and hasattr(file_obj, "tell"):
            current_pos = file_obj.tell()
            file_obj.seek(0, 2)  # Seek to end
            file_size = file_obj.tell()
            file_obj.seek(current_pos)  # Restore position
        elif hasattr(file_obj, "content_length") and file_obj.content_length:
            file_size = file_obj.content_length
        else:
            return True  # Can't determine size, allow it

        max_size_bytes = max_size_mb * 1024 * 1024
        return file_size <= max_size_bytes
    except Exception as e:
        logger.warning(f"Could not validate file size: {e}")
        return True  # Allow if we can't check


def create_temp_directory(prefix: str = "excel_pptx_merger_") -> str:
    """Create a temporary directory with the given prefix.

    This function is maintained for backward compatibility.
    New code should use the storage abstraction layer instead.

    Args:
        prefix: Prefix for the temporary directory name

    Returns:
        Path to the created temporary directory
    """
    temp_dir = tempfile.mkdtemp(prefix=prefix)
    logger.debug(f"Created temporary directory: {temp_dir}")
    return temp_dir


def save_uploaded_file(
    file_obj: Union[BinaryIO, FileStorage],
    destination_path: str,
    allowed_extensions: Optional[List[str]] = None,
    max_size_mb: Optional[int] = None,
) -> str:
    """Save uploaded file to destination with validation."""
    try:
        # Get filename
        filename = getattr(file_obj, "filename", "uploaded_file")
        if not filename:
            filename = "uploaded_file"

        # Validate extension if specified
        if allowed_extensions and not validate_file_extension(
            filename, allowed_extensions
        ):
            raise ValidationError(
                f"File extension not allowed. Allowed: {', '.join(allowed_extensions)}"
            )

        # Validate file size if specified
        if max_size_mb and not validate_file_size(file_obj, max_size_mb):
            raise ValidationError(
                f"File size exceeds maximum allowed size of {max_size_mb}MB"
            )

        # Ensure destination directory exists
        os.makedirs(os.path.dirname(destination_path), exist_ok=True)

        # Save file
        if hasattr(file_obj, "save"):
            # Flask FileStorage object
            file_obj.save(destination_path)
        else:
            # Regular file-like object
            with open(destination_path, "wb") as dest_file:
                if hasattr(file_obj, "read"):
                    shutil.copyfileobj(file_obj, dest_file)
                else:
                    dest_file.write(file_obj)

        logger.info(f"Saved uploaded file to: {destination_path}")
        return destination_path

    except (ValidationError, FileProcessingError):
        raise
    except Exception as e:
        raise FileProcessingError(f"Failed to save uploaded file: {e}")


def cleanup_directory(directory_path: str, force: bool = False) -> None:
    """Clean up a directory and all its contents.

    This function is maintained for backward compatibility.
    New code should use the storage abstraction layer instead.

    Args:
        directory_path: Path to the directory to clean up
        force: If True, ignore errors during cleanup
    """
    try:
        if os.path.exists(directory_path):
            shutil.rmtree(directory_path, ignore_errors=force)
            logger.debug(f"Cleaned up directory: {directory_path}")
        else:
            logger.debug(
                f"Directory does not exist, skipping cleanup: {directory_path}"
            )
    except Exception as e:
        if not force:
            raise
        logger.warning(f"Error during directory cleanup (ignored): {e}")


def get_file_info(file_path: str) -> dict:
    """Get information about a file."""
    try:
        path_obj = Path(file_path)
        if not path_obj.exists():
            raise FileProcessingError(f"File does not exist: {file_path}")

        stat = path_obj.stat()
        return {
            "name": path_obj.name,
            "size": stat.st_size,
            "extension": path_obj.suffix.lower().lstrip("."),
            "modified": stat.st_mtime,
            "absolute_path": str(path_obj.absolute()),
        }
    except FileProcessingError:
        raise
    except Exception as e:
        raise FileProcessingError(f"Failed to get file info for {file_path}: {e}")


def ensure_directory_exists(directory_path: str) -> None:
    """Ensure directory exists, create if it doesn't."""
    try:
        os.makedirs(directory_path, exist_ok=True)
    except Exception as e:
        raise FileProcessingError(f"Failed to create directory {directory_path}: {e}")


def copy_file(source_path: str, destination_path: str) -> None:
    """Copy file from source to destination."""
    try:
        ensure_directory_exists(os.path.dirname(destination_path))
        shutil.copy2(source_path, destination_path)
        logger.debug(f"Copied file from {source_path} to {destination_path}")
    except Exception as e:
        raise FileProcessingError(f"Failed to copy file: {e}")


def generate_unique_filename(base_path: str, extension: str) -> str:
    """Generate unique filename to avoid conflicts."""
    counter = 1
    original_path = f"{base_path}.{extension.lstrip('.')}"

    if not os.path.exists(original_path):
        return original_path

    while True:
        new_path = f"{base_path}_{counter}.{extension.lstrip('.')}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1
