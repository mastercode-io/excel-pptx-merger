"""
Job Queue System - General purpose async job processing

This module provides a flexible job queue system that can make any existing
endpoint asynchronous without modifying the original endpoint implementations.
"""

import uuid
import threading
import time
from datetime import datetime, timezone
from typing import Dict, Any, Optional, Callable
from dataclasses import dataclass, asdict
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class JobStatus(Enum):
    """Job status enumeration"""
    PENDING = "pending"
    RUNNING = "running" 
    COMPLETED = "completed"
    FAILED = "failed"
    EXPIRED = "expired"


@dataclass
class Job:
    """Job data structure"""
    id: str
    endpoint: str
    payload: Dict[str, Any]
    status: JobStatus
    progress: int = 0
    created_at: str = ""
    updated_at: str = ""
    result: Optional[Any] = None
    error: Optional[str] = None
    retry_count: int = 0
    
    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.now(timezone.utc).isoformat()
        if not self.updated_at:
            self.updated_at = self.created_at
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert job to dictionary, handling enum serialization"""
        data = asdict(self)
        data['status'] = self.status.value
        return data
    
    def update_status(self, status: JobStatus, progress: int = None, error: str = None):
        """Update job status with timestamp"""
        self.status = status
        if progress is not None:
            self.progress = progress
        if error is not None:
            self.error = error
        self.updated_at = datetime.now(timezone.utc).isoformat()


class JobQueue:
    """
    General purpose job queue system
    
    Provides async processing for any endpoint by calling handler functions
    directly instead of making HTTP requests.
    """
    
    def __init__(self):
        self._jobs: Dict[str, Job] = {}
        self._lock = threading.RLock()
        self._max_jobs_per_client = 10
        self._job_timeout_seconds = 300  # 5 minutes
        self._cleanup_interval = 3600  # 1 hour
        self._last_cleanup = time.time()
        
        # Allowed endpoints for async processing
        self._allowed_endpoints = {
            '/api/v1/extract',
            '/api/v1/merge', 
            '/api/v1/update'
        }
    
    def generate_job_id(self) -> str:
        """Generate unique job ID"""
        timestamp = int(time.time() * 1000)
        unique_id = str(uuid.uuid4())[:8]
        return f"job_{timestamp}_{unique_id}"
    
    def validate_endpoint(self, endpoint: str) -> bool:
        """Validate that endpoint is allowed for async processing"""
        return endpoint in self._allowed_endpoints
    
    def create_job(self, endpoint: str, payload: Dict[str, Any], client_ip: str = None) -> str:
        """
        Create a new job
        
        Args:
            endpoint: Target endpoint (e.g., '/api/v1/extract')
            payload: Request payload for the endpoint
            client_ip: Client IP for rate limiting (optional)
            
        Returns:
            Job ID
            
        Raises:
            ValueError: If endpoint is not allowed or rate limit exceeded
        """
        with self._lock:
            # Validate endpoint
            if not self.validate_endpoint(endpoint):
                raise ValueError(f"Endpoint '{endpoint}' is not allowed for async processing")
            
            # Rate limiting check (if client_ip provided)
            if client_ip:
                active_jobs = sum(1 for job in self._jobs.values() 
                                if job.status in [JobStatus.PENDING, JobStatus.RUNNING])
                if active_jobs >= self._max_jobs_per_client:
                    raise ValueError(f"Too many active jobs. Maximum {self._max_jobs_per_client} allowed.")
            
            # Create job
            job_id = self.generate_job_id()
            job = Job(
                id=job_id,
                endpoint=endpoint,
                payload=payload,
                status=JobStatus.PENDING
            )
            
            self._jobs[job_id] = job
            
            logger.info(f"Created job {job_id} for endpoint {endpoint}")
            return job_id
    
    def get_job(self, job_id: str) -> Optional[Job]:
        """Get job by ID"""
        with self._lock:
            return self._jobs.get(job_id)
    
    def get_job_status(self, job_id: str) -> Optional[Dict[str, Any]]:
        """
        Get job status information
        
        Returns:
            Job status dict or None if job not found
        """
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return None
                
            return {
                'success': True,
                'jobId': job.id,
                'status': job.status.value,
                'progress': job.progress,
                'message': self._get_status_message(job),
                'created_at': job.created_at,
                'updated_at': job.updated_at
            }
    
    def get_job_result(self, job_id: str, cleanup: bool = True) -> Optional[Dict[str, Any]]:
        """
        Get job result and optionally clean up storage
        
        Args:
            job_id: Job ID
            cleanup: Whether to delete job after retrieving result
            
        Returns:
            Job result dict or None if job not found/not ready
        """
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return None
            
            if job.status == JobStatus.COMPLETED:
                result = {
                    'success': True,
                    'jobId': job.id,
                    'status': job.status.value,
                    'data': job.result,
                    'retrieved_at': datetime.now(timezone.utc).isoformat()
                }
                
                # Cleanup job if requested
                if cleanup:
                    del self._jobs[job_id]
                    logger.info(f"Cleaned up completed job {job_id}")
                
                return result
            
            elif job.status == JobStatus.FAILED:
                result = {
                    'success': False,
                    'jobId': job.id,
                    'status': job.status.value,
                    'error': job.error,
                    'retrieved_at': datetime.now(timezone.utc).isoformat()
                }
                
                # Cleanup failed job if requested
                if cleanup:
                    del self._jobs[job_id]
                    logger.info(f"Cleaned up failed job {job_id}")
                
                return result
            
            else:
                # Job not ready
                return {
                    'success': False,
                    'jobId': job.id,
                    'status': job.status.value,
                    'error': 'Job still processing',
                    'message': 'Use /jobs/{jobId}/status to check progress'
                }
    
    def update_job_progress(self, job_id: str, progress: int, message: str = None):
        """Update job progress"""
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.progress = progress
                job.updated_at = datetime.now(timezone.utc).isoformat()
                if message:
                    logger.info(f"Job {job_id} progress: {progress}% - {message}")
    
    def complete_job(self, job_id: str, result: Any):
        """Mark job as completed with result"""
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.update_status(JobStatus.COMPLETED, progress=100)
                job.result = result
                logger.info(f"Job {job_id} completed successfully")
    
    def fail_job(self, job_id: str, error: str):
        """Mark job as failed with error message"""
        with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.update_status(JobStatus.FAILED, error=error)
                logger.error(f"Job {job_id} failed: {error}")
    
    def process_job(self, job_id: str, handler_func: Callable):
        """
        Process a job using the provided handler function
        
        Args:
            job_id: Job ID to process
            handler_func: Function to call with job payload
        """
        job = self.get_job(job_id)
        if not job:
            logger.error(f"Job {job_id} not found for processing")
            return
        
        try:
            # Update status to running
            with self._lock:
                job.update_status(JobStatus.RUNNING, progress=10)
            
            logger.info(f"Processing job {job_id} for endpoint {job.endpoint}")
            
            # Call handler function with payload
            result = handler_func(job.payload)
            
            # Complete job
            self.complete_job(job_id, result)
            
        except Exception as e:
            error_msg = str(e)
            logger.exception(f"Job {job_id} processing failed: {error_msg}")
            self.fail_job(job_id, error_msg)
    
    def list_jobs(self, status_filter: str = None) -> Dict[str, Any]:
        """
        List all jobs, optionally filtered by status
        
        Args:
            status_filter: Optional status to filter by
            
        Returns:
            Dict with jobs list and metadata
        """
        with self._lock:
            jobs = list(self._jobs.values())
            
            if status_filter:
                try:
                    status_enum = JobStatus(status_filter)
                    jobs = [job for job in jobs if job.status == status_enum]
                except ValueError:
                    # Invalid status filter
                    pass
            
            return {
                'success': True,
                'total_jobs': len(jobs),
                'jobs': [job.to_dict() for job in jobs]
            }
    
    def delete_job(self, job_id: str) -> bool:
        """
        Delete/cancel a job
        
        Args:
            job_id: Job ID to delete
            
        Returns:
            True if job was deleted, False if not found
        """
        with self._lock:
            if job_id in self._jobs:
                del self._jobs[job_id]
                logger.info(f"Deleted job {job_id}")
                return True
            return False
    
    def cleanup_expired_jobs(self):
        """Clean up expired and old jobs"""
        current_time = time.time()
        
        # Only run cleanup if enough time has passed
        if current_time - self._last_cleanup < self._cleanup_interval:
            return
        
        with self._lock:
            expired_jobs = []
            current_datetime = datetime.now(timezone.utc)
            
            for job_id, job in self._jobs.items():
                created_at = datetime.fromisoformat(job.created_at.replace('Z', '+00:00'))
                age_seconds = (current_datetime - created_at).total_seconds()
                
                # Mark as expired if too old
                if age_seconds > self._job_timeout_seconds:
                    if job.status in [JobStatus.PENDING, JobStatus.RUNNING]:
                        job.update_status(JobStatus.EXPIRED, error="Job expired due to timeout")
                    
                    # Remove very old jobs (completed, failed, expired)
                    if age_seconds > self._job_timeout_seconds * 2:
                        expired_jobs.append(job_id)
            
            # Remove expired jobs
            for job_id in expired_jobs:
                del self._jobs[job_id]
                logger.info(f"Cleaned up expired job {job_id}")
            
            self._last_cleanup = current_time
            
            if expired_jobs:
                logger.info(f"Cleaned up {len(expired_jobs)} expired jobs")
    
    def _get_status_message(self, job: Job) -> str:
        """Get human-readable status message"""
        if job.status == JobStatus.PENDING:
            return "Job queued for processing"
        elif job.status == JobStatus.RUNNING:
            return f"Processing {job.endpoint} request..."
        elif job.status == JobStatus.COMPLETED:
            return "Job completed successfully"
        elif job.status == JobStatus.FAILED:
            return f"Job failed: {job.error or 'Unknown error'}"
        elif job.status == JobStatus.EXPIRED:
            return "Job expired due to timeout"
        else:
            return "Unknown status"
    
    def get_stats(self) -> Dict[str, Any]:
        """Get job queue statistics"""
        with self._lock:
            stats = {
                'total_jobs': len(self._jobs),
                'by_status': {},
                'by_endpoint': {}
            }
            
            for job in self._jobs.values():
                # Count by status
                status = job.status.value
                stats['by_status'][status] = stats['by_status'].get(status, 0) + 1
                
                # Count by endpoint
                endpoint = job.endpoint
                stats['by_endpoint'][endpoint] = stats['by_endpoint'].get(endpoint, 0) + 1
            
            return stats


# Global job queue instance
job_queue = JobQueue()