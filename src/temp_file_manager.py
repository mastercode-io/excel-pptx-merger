"""Temporary file management with environment-based cleanup control and cloud storage support."""

import os
import shutil
import tempfile
import threading
import time
from contextlib import contextmanager
from typing import Dict, List, Optional, Any, Union, BinaryIO
import logging
import atexit
from datetime import datetime, timedelta
import uuid

from .utils.exceptions import TempFileError
from .utils.file_utils import create_temp_directory, cleanup_directory
from .utils.storage import StorageFactory, StorageBackend

logger = logging.getLogger(__name__)


class TempFileManager:
    """Manages temporary files and directories with configurable cleanup policies and storage backends."""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None) -> None:
        """Initialize temporary file manager with configuration."""
        self.config = config or {}
        self._temp_directories: Dict[str, Dict[str, Any]] = {}
        self._cleanup_threads: List[threading.Thread] = []
        self._lock = threading.Lock()
        
        # Register cleanup on exit
        atexit.register(self._cleanup_on_exit)
        
        # Extract configuration
        self.enabled = self._get_config_bool('enabled', True)
        self.delay_seconds = self._get_config_int('delay_seconds', 300)
        self.keep_on_error = self._get_config_bool('keep_on_error', True)
        self.development_mode = self._get_config_bool('development_mode', False)
        
        # Initialize storage backend
        self.storage = StorageFactory.get_storage_backend()
        storage_type = os.getenv('STORAGE_BACKEND', 'LOCAL').upper()
        
        logger.debug(f"TempFileManager initialized: enabled={self.enabled}, "
                    f"delay={self.delay_seconds}s, dev_mode={self.development_mode}, "
                    f"storage_backend={storage_type}")
    
    def _get_config_bool(self, key: str, default: bool) -> bool:
        """Get boolean configuration value with environment override."""
        # Check environment variables first
        env_key = f"TEMP_FILE_{key.upper()}"
        env_value = os.getenv(env_key)
        if env_value is not None:
            return env_value.lower() in ('true', '1', 'yes', 'on')
        
        # Then check config dict
        return self.config.get(key, default)
    
    def _get_config_int(self, key: str, default: int) -> int:
        """Get integer configuration value with environment override."""
        # Check environment variables first
        env_key = f"TEMP_FILE_{key.upper()}"
        env_value = os.getenv(env_key)
        if env_value is not None:
            try:
                return int(env_value)
            except ValueError:
                pass
        
        # Then check config dict
        return self.config.get(key, default)
    
    def create_temp_directory(self, prefix: str = "excel_pptx_merger_", 
                            cleanup_delay: Optional[int] = None,
                            keep_on_error: Optional[bool] = None) -> str:
        """Create a temporary directory with configured cleanup policy."""
        try:
            # Generate a unique directory path
            dir_id = str(uuid.uuid4())
            temp_dir = f"{prefix}{dir_id}"
            
            # Create the directory in the storage backend
            storage_path = self.storage.create_directory(temp_dir)
            
            # Store directory info for cleanup
            with self._lock:
                self._temp_directories[temp_dir] = {
                    'created_at': datetime.now(),
                    'cleanup_delay': cleanup_delay or self.delay_seconds,
                    'keep_on_error': keep_on_error if keep_on_error is not None else self.keep_on_error,
                    'error_occurred': False,
                    'manual_cleanup': False,
                    'cleanup_scheduled': False,
                    'storage_path': storage_path
                }
            
            logger.debug(f"Created temporary directory: {temp_dir}")
            return temp_dir
        
        except Exception as e:
            raise TempFileError(f"Failed to create temporary directory: {e}")
    
    @contextmanager
    def temp_directory(self, prefix: str = "excel_pptx_merger_",
                      cleanup_delay: Optional[int] = None,
                      keep_on_error: Optional[bool] = None):
        """Context manager for temporary directory with automatic cleanup."""
        temp_dir = None
        try:
            temp_dir = self.create_temp_directory(prefix, cleanup_delay, keep_on_error)
            yield temp_dir
        except Exception as e:
            if temp_dir:
                self.mark_error(temp_dir)
            raise
        finally:
            if temp_dir:
                self.schedule_cleanup(temp_dir)
    
    def mark_error(self, temp_dir: str) -> None:
        """Mark a temporary directory as having an error."""
        with self._lock:
            if temp_dir in self._temp_directories:
                self._temp_directories[temp_dir]['error_occurred'] = True
                logger.debug(f"Marked directory as error: {temp_dir}")
    
    def schedule_cleanup(self, temp_dir: str, delay_override: Optional[int] = None) -> None:
        """Schedule cleanup for a temporary directory."""
        if not self.enabled:
            logger.debug(f"Cleanup disabled, skipping: {temp_dir}")
            return
        
        if self.development_mode:
            logger.debug(f"Development mode, skipping cleanup: {temp_dir}")
            return
        
        with self._lock:
            if temp_dir not in self._temp_directories:
                logger.warning(f"Directory not tracked, cannot schedule cleanup: {temp_dir}")
                return
            
            dir_info = self._temp_directories[temp_dir]
            
            # Check if we should keep on error
            if dir_info['error_occurred'] and dir_info['keep_on_error']:
                logger.info(f"Keeping directory due to error: {temp_dir}")
                return
            
            # Prevent double scheduling
            if dir_info['cleanup_scheduled']:
                logger.debug(f"Cleanup already scheduled for: {temp_dir}")
                return
            
            dir_info['cleanup_scheduled'] = True
        
        # Use override delay or configured delay
        delay = delay_override or self._temp_directories[temp_dir]['cleanup_delay']
        
        if delay <= 0:
            # Immediate cleanup
            self._cleanup_directory(temp_dir)
        else:
            # Scheduled cleanup
            cleanup_thread = threading.Thread(
                target=self._delayed_cleanup,
                args=(temp_dir, delay),
                daemon=True
            )
            cleanup_thread.start()
            self._cleanup_threads.append(cleanup_thread)
            
            logger.debug(f"Scheduled cleanup for {temp_dir} in {delay} seconds")
    
    def _delayed_cleanup(self, temp_dir: str, delay: int) -> None:
        """Perform delayed cleanup of temporary directory."""
        try:
            time.sleep(delay)
            self._cleanup_directory(temp_dir)
        except Exception as e:
            logger.error(f"Failed delayed cleanup for {temp_dir}: {e}")
    
    def _cleanup_directory(self, temp_dir: str) -> None:
        """Perform actual cleanup of temporary directory."""
        try:
            with self._lock:
                if temp_dir in self._temp_directories:
                    # Delete from storage backend
                    self.storage.delete_directory(temp_dir)
                    # Remove from tracking
                    del self._temp_directories[temp_dir]
                    logger.info(f"Cleaned up temporary directory: {temp_dir}")
                else:
                    logger.debug(f"Directory not tracked, skipping cleanup: {temp_dir}")
        
        except Exception as e:
            logger.error(f"Failed to cleanup directory {temp_dir}: {e}")
    
    def cleanup_immediately(self, temp_dir: str) -> None:
        """Immediately cleanup a temporary directory."""
        with self._lock:
            if temp_dir in self._temp_directories:
                self._temp_directories[temp_dir]['manual_cleanup'] = True
        
        self._cleanup_directory(temp_dir)
    
    def cleanup_all(self, force: bool = False) -> None:
        """Cleanup all tracked temporary directories."""
        directories_to_cleanup = []
        
        with self._lock:
            for temp_dir, dir_info in list(self._temp_directories.items()):
                if force or not (dir_info['error_occurred'] and dir_info['keep_on_error']):
                    directories_to_cleanup.append(temp_dir)
        
        for temp_dir in directories_to_cleanup:
            self._cleanup_directory(temp_dir)
        
        logger.info(f"Cleaned up {len(directories_to_cleanup)} directories")
    
    def list_temp_directories(self) -> List[Dict[str, Any]]:
        """List all tracked temporary directories with their info."""
        with self._lock:
            return [
                {
                    'path': temp_dir,
                    'created_at': info['created_at'].isoformat(),
                    'age_seconds': (datetime.now() - info['created_at']).total_seconds(),
                    'cleanup_delay': info['cleanup_delay'],
                    'keep_on_error': info['keep_on_error'],
                    'error_occurred': info['error_occurred'],
                    'cleanup_scheduled': info['cleanup_scheduled'],
                    'storage_path': info.get('storage_path', '')
                }
                for temp_dir, info in self._temp_directories.items()
            ]
    
    def cleanup_old_directories(self, max_age_hours: int = 24) -> None:
        """Cleanup directories older than specified age."""
        cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
        directories_to_cleanup = []
        
        with self._lock:
            for temp_dir, dir_info in list(self._temp_directories.items()):
                if dir_info['created_at'] < cutoff_time:
                    # Skip if error occurred and we should keep on error
                    if dir_info['error_occurred'] and dir_info['keep_on_error']:
                        continue
                    directories_to_cleanup.append(temp_dir)
        
        for temp_dir in directories_to_cleanup:
            self._cleanup_directory(temp_dir)
        
        if directories_to_cleanup:
            logger.info(f"Cleaned up {len(directories_to_cleanup)} old directories")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get statistics about temporary file management."""
        with self._lock:
            total_dirs = len(self._temp_directories)
            error_dirs = sum(1 for info in self._temp_directories.values() if info['error_occurred'])
            scheduled_dirs = sum(1 for info in self._temp_directories.values() if info['cleanup_scheduled'])
        
        storage_type = os.getenv('STORAGE_BACKEND', 'LOCAL').upper()
        
        return {
            'total_directories': total_dirs,
            'directories_with_errors': error_dirs,
            'scheduled_for_cleanup': scheduled_dirs,
            'cleanup_enabled': self.enabled,
            'development_mode': self.development_mode,
            'default_delay_seconds': self.delay_seconds,
            'active_cleanup_threads': len([t for t in self._cleanup_threads if t.is_alive()]),
            'storage_backend': storage_type
        }
    
    def set_config(self, config: Dict[str, Any]) -> None:
        """Update configuration at runtime."""
        self.config = config
        self.enabled = self._get_config_bool('enabled', True)
        self.delay_seconds = self._get_config_int('delay_seconds', 300)
        self.keep_on_error = self._get_config_bool('keep_on_error', True)
        self.development_mode = self._get_config_bool('development_mode', False)
        
        logger.info("TempFileManager configuration updated")
    
    def enable_development_mode(self) -> None:
        """Enable development mode (disable cleanup)."""
        self.development_mode = True
        logger.info("Development mode enabled - cleanup disabled")
    
    def disable_development_mode(self) -> None:
        """Disable development mode (enable cleanup)."""
        self.development_mode = False
        logger.info("Development mode disabled - cleanup enabled")
    
    def _cleanup_on_exit(self) -> None:
        """Cleanup all temporary directories on application exit."""
        if self.development_mode:
            logger.debug("Development mode - skipping exit cleanup")
            return
        
        logger.info("Application exit - cleaning up temporary directories")
        self.cleanup_all(force=True)
        
        # Wait for cleanup threads to finish
        for thread in self._cleanup_threads:
            if thread.is_alive():
                thread.join(timeout=5)
    
    def create_temp_file(self, temp_dir: str, filename: str, content: Union[bytes, BinaryIO] = b"") -> str:
        """Create a temporary file within a managed directory."""
        try:
            if temp_dir not in self._temp_directories:
                raise TempFileError(f"Directory not managed by TempFileManager: {temp_dir}")
            
            file_path = f"{temp_dir}/{filename}"
            
            # Save to storage backend
            storage_path = self.storage.save_file(file_path, content)
            
            logger.debug(f"Created temporary file: {file_path}")
            return file_path
        
        except Exception as e:
            raise TempFileError(f"Failed to create temporary file: {e}")
    
    def save_file_to_temp(self, temp_dir: str, filename: str, file_obj) -> str:
        """Save uploaded file to temporary directory."""
        try:
            if temp_dir not in self._temp_directories:
                raise TempFileError(f"Directory not managed by TempFileManager: {temp_dir}")
            
            file_path = f"{temp_dir}/{filename}"
            
            # Save to storage backend
            if hasattr(file_obj, 'read'):
                storage_path = self.storage.save_file(file_path, file_obj)
            else:
                storage_path = self.storage.save_file(file_path, file_obj)
            
            logger.debug(f"Saved file to temporary directory: {file_path}")
            # Return the full storage path instead of the relative path
            return storage_path
        
        except Exception as e:
            raise TempFileError(f"Failed to save file to temporary directory: {e}")
    
    def get_file_content(self, file_path: str) -> bytes:
        """Get content of a file from storage."""
        try:
            return self.storage.read_file(file_path)
        except Exception as e:
            raise TempFileError(f"Failed to read file content: {e}")
    
    def get_public_url(self, file_path: str, expiration_seconds: int = 3600) -> Optional[str]:
        """Get a public URL for accessing the file."""
        try:
            return self.storage.get_public_url(file_path, expiration_seconds)
        except Exception as e:
            logger.error(f"Failed to get public URL for {file_path}: {e}")
            return None
    
    def __del__(self) -> None:
        """Cleanup on object destruction."""
        try:
            if hasattr(self, '_temp_directories') and not self.development_mode:
                self.cleanup_all(force=True)
        except Exception:
            pass  # Ignore errors during destruction