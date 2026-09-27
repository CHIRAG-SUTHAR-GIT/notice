"""
Session Manager for the Fraud Analysis Application.

Handles user sessions and data lifecycle with in-memory storage.
Ensures no data is stored on server after session ends.
"""

import uuid
from datetime import datetime, timedelta
from typing import Any, Dict, Optional

from src.models import SessionInfo


class SessionManager:
    """
    Manages user sessions with in-memory storage and automatic timeout.
    
    Attributes:
        SESSION_TIMEOUT_MINUTES: Session timeout duration (30 minutes)
    """
    
    SESSION_TIMEOUT_MINUTES: int = 30
    
    def __init__(self):
        """Initialize the session manager with empty storage."""
        self._sessions: Dict[str, SessionInfo] = {}
        self._session_data: Dict[str, Dict[str, Any]] = {}
    
    def create_session(self, user_id: Optional[str] = None) -> str:
        """
        Create a new session with a unique ID.
        
        Args:
            user_id: Optional user identifier
            
        Returns:
            Unique session ID string
        """
        session_id = str(uuid.uuid4())
        now = datetime.now()
        
        session_info = SessionInfo(
            session_id=session_id,
            created_at=now,
            last_activity=now,
            user_id=user_id
        )
        
        self._sessions[session_id] = session_info
        self._session_data[session_id] = {}
        
        return session_id
    
    def validate_session(self, session_id: str) -> bool:
        """
        Check if a session is valid and not expired.
        
        Args:
            session_id: The session ID to validate
            
        Returns:
            True if session is valid and not timed out, False otherwise
        """
        if session_id not in self._sessions:
            return False
        
        if self.check_timeout(session_id):
            # Session has timed out, clean it up
            self.cleanup_session(session_id)
            return False
        
        # Update last activity timestamp
        self._sessions[session_id].last_activity = datetime.now()
        return True
    
    def store_data(self, session_id: str, key: str, data: Any) -> None:
        """
        Store data in session (in-memory only).
        
        Args:
            session_id: The session ID
            key: Key to store data under
            data: Data to store
            
        Raises:
            ValueError: If session does not exist or is invalid
        """
        if not self.validate_session(session_id):
            raise ValueError(f"Invalid or expired session: {session_id}")
        
        self._session_data[session_id][key] = data
    
    def get_data(self, session_id: str, key: str) -> Any:
        """
        Retrieve data from session.
        
        Args:
            session_id: The session ID
            key: Key to retrieve data for
            
        Returns:
            Stored data or None if key doesn't exist
            
        Raises:
            ValueError: If session does not exist or is invalid
        """
        if not self.validate_session(session_id):
            raise ValueError(f"Invalid or expired session: {session_id}")
        
        return self._session_data[session_id].get(key)
    
    def cleanup_session(self, session_id: str) -> None:
        """
        Delete all session data permanently.
        
        Ensures no data remains on server after session ends.
        
        Args:
            session_id: The session ID to clean up
        """
        # Remove session info
        if session_id in self._sessions:
            del self._sessions[session_id]
        
        # Remove all session data
        if session_id in self._session_data:
            # Clear all data in the session
            self._session_data[session_id].clear()
            del self._session_data[session_id]
    
    def check_timeout(self, session_id: str) -> bool:
        """
        Check if a session has timed out (30 minutes of inactivity).
        
        Args:
            session_id: The session ID to check
            
        Returns:
            True if session has timed out, False otherwise
        """
        if session_id not in self._sessions:
            return True
        
        session = self._sessions[session_id]
        timeout_threshold = timedelta(minutes=self.SESSION_TIMEOUT_MINUTES)
        time_since_activity = datetime.now() - session.last_activity
        
        return time_since_activity > timeout_threshold
    
    def get_session_info(self, session_id: str) -> Optional[SessionInfo]:
        """
        Get session information.
        
        Args:
            session_id: The session ID
            
        Returns:
            SessionInfo object or None if session doesn't exist
        """
        if not self.validate_session(session_id):
            return None
        
        return self._sessions.get(session_id)
    
    def set_input_filename(self, session_id: str, filename: str) -> None:
        """
        Set the input filename for a session.
        
        Args:
            session_id: The session ID
            filename: The input filename
            
        Raises:
            ValueError: If session does not exist or is invalid
        """
        if not self.validate_session(session_id):
            raise ValueError(f"Invalid or expired session: {session_id}")
        
        self._sessions[session_id].input_filename = filename
    
    def cleanup_expired_sessions(self) -> int:
        """
        Clean up all expired sessions.
        
        Returns:
            Number of sessions cleaned up
        """
        expired_sessions = [
            session_id for session_id in list(self._sessions.keys())
            if self.check_timeout(session_id)
        ]
        
        for session_id in expired_sessions:
            self.cleanup_session(session_id)
        
        return len(expired_sessions)
