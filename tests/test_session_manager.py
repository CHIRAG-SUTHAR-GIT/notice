"""
Tests for the Session Manager module.

Tests session creation, validation, data storage, cleanup, and timeout functionality.
"""

import time
from datetime import datetime, timedelta
from unittest.mock import patch

import pytest

from src.session_manager import SessionManager


class TestSessionManager:
    """Tests for SessionManager class."""
    
    def test_create_session_returns_unique_id(self):
        """Test that create_session returns a unique session ID."""
        manager = SessionManager()
        session1 = manager.create_session()
        session2 = manager.create_session()
        
        assert session1 is not None
        assert session2 is not None
        assert session1 != session2
    
    def test_create_session_with_user_id(self):
        """Test that create_session stores user_id."""
        manager = SessionManager()
        session_id = manager.create_session(user_id="test_user")
        
        session_info = manager.get_session_info(session_id)
        assert session_info is not None
        assert session_info.user_id == "test_user"
    
    def test_validate_session_valid(self):
        """Test that validate_session returns True for valid session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        assert manager.validate_session(session_id) is True
    
    def test_validate_session_invalid(self):
        """Test that validate_session returns False for invalid session."""
        manager = SessionManager()
        
        assert manager.validate_session("invalid-session-id") is False
    
    def test_store_and_get_data(self):
        """Test storing and retrieving data from session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        test_data = {"key": "value", "number": 42}
        manager.store_data(session_id, "test_key", test_data)
        
        retrieved = manager.get_data(session_id, "test_key")
        assert retrieved == test_data
    
    def test_get_data_nonexistent_key(self):
        """Test that get_data returns None for nonexistent key."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        result = manager.get_data(session_id, "nonexistent_key")
        assert result is None
    
    def test_store_data_invalid_session(self):
        """Test that store_data raises error for invalid session."""
        manager = SessionManager()
        
        with pytest.raises(ValueError, match="Invalid or expired session"):
            manager.store_data("invalid-session", "key", "value")
    
    def test_get_data_invalid_session(self):
        """Test that get_data raises error for invalid session."""
        manager = SessionManager()
        
        with pytest.raises(ValueError, match="Invalid or expired session"):
            manager.get_data("invalid-session", "key")
    
    def test_cleanup_session(self):
        """Test that cleanup_session removes all session data."""
        manager = SessionManager()
        session_id = manager.create_session()
        manager.store_data(session_id, "key", "value")
        
        manager.cleanup_session(session_id)
        
        assert manager.validate_session(session_id) is False
    
    def test_cleanup_nonexistent_session(self):
        """Test that cleanup_session handles nonexistent session gracefully."""
        manager = SessionManager()
        
        # Should not raise an error
        manager.cleanup_session("nonexistent-session")
    
    def test_check_timeout_not_expired(self):
        """Test that check_timeout returns False for active session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        assert manager.check_timeout(session_id) is False
    
    def test_check_timeout_expired(self):
        """Test that check_timeout returns True for expired session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        # Manually set last_activity to 31 minutes ago
        manager._sessions[session_id].last_activity = (
            datetime.now() - timedelta(minutes=31)
        )
        
        assert manager.check_timeout(session_id) is True
    
    def test_check_timeout_nonexistent_session(self):
        """Test that check_timeout returns True for nonexistent session."""
        manager = SessionManager()
        
        assert manager.check_timeout("nonexistent-session") is True
    
    def test_validate_session_cleans_up_expired(self):
        """Test that validate_session cleans up expired sessions."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        # Manually set last_activity to 31 minutes ago
        manager._sessions[session_id].last_activity = (
            datetime.now() - timedelta(minutes=31)
        )
        
        # Validate should return False and clean up
        assert manager.validate_session(session_id) is False
        
        # Session should be completely removed
        assert session_id not in manager._sessions
        assert session_id not in manager._session_data
    
    def test_validate_session_updates_last_activity(self):
        """Test that validate_session updates last_activity timestamp."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        # Set last_activity to 10 minutes ago
        old_time = datetime.now() - timedelta(minutes=10)
        manager._sessions[session_id].last_activity = old_time
        
        # Validate session
        manager.validate_session(session_id)
        
        # Last activity should be updated to now
        new_time = manager._sessions[session_id].last_activity
        assert new_time > old_time
    
    def test_set_input_filename(self):
        """Test setting input filename for session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        manager.set_input_filename(session_id, "test_file.xlsx")
        
        session_info = manager.get_session_info(session_id)
        assert session_info.input_filename == "test_file.xlsx"
    
    def test_set_input_filename_invalid_session(self):
        """Test that set_input_filename raises error for invalid session."""
        manager = SessionManager()
        
        with pytest.raises(ValueError, match="Invalid or expired session"):
            manager.set_input_filename("invalid-session", "test.xlsx")
    
    def test_get_session_info_valid(self):
        """Test getting session info for valid session."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        session_info = manager.get_session_info(session_id)
        
        assert session_info is not None
        assert session_info.session_id == session_id
        assert session_info.created_at is not None
        assert session_info.last_activity is not None
    
    def test_get_session_info_invalid(self):
        """Test that get_session_info returns None for invalid session."""
        manager = SessionManager()
        
        result = manager.get_session_info("invalid-session")
        assert result is None
    
    def test_cleanup_expired_sessions(self):
        """Test cleaning up multiple expired sessions."""
        manager = SessionManager()
        
        # Create some sessions
        session1 = manager.create_session()
        session2 = manager.create_session()
        session3 = manager.create_session()
        
        # Expire session1 and session2
        manager._sessions[session1].last_activity = (
            datetime.now() - timedelta(minutes=31)
        )
        manager._sessions[session2].last_activity = (
            datetime.now() - timedelta(minutes=45)
        )
        
        # Clean up expired sessions
        cleaned = manager.cleanup_expired_sessions()
        
        assert cleaned == 2
        assert session1 not in manager._sessions
        assert session2 not in manager._sessions
        assert session3 in manager._sessions
    
    def test_session_timeout_constant(self):
        """Test that session timeout is 30 minutes."""
        assert SessionManager.SESSION_TIMEOUT_MINUTES == 30
    
    def test_in_memory_storage_isolation(self):
        """Test that sessions have isolated data storage."""
        manager = SessionManager()
        session1 = manager.create_session()
        session2 = manager.create_session()
        
        manager.store_data(session1, "key", "value1")
        manager.store_data(session2, "key", "value2")
        
        assert manager.get_data(session1, "key") == "value1"
        assert manager.get_data(session2, "key") == "value2"
    
    def test_data_cleared_on_cleanup(self):
        """Test that all data is cleared when session is cleaned up."""
        manager = SessionManager()
        session_id = manager.create_session()
        
        # Store multiple pieces of data
        manager.store_data(session_id, "key1", "value1")
        manager.store_data(session_id, "key2", {"nested": "data"})
        manager.store_data(session_id, "key3", [1, 2, 3])
        
        # Cleanup
        manager.cleanup_session(session_id)
        
        # Verify session is completely gone
        assert session_id not in manager._sessions
        assert session_id not in manager._session_data
