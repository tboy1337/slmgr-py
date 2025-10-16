"""Pytest configuration and fixtures"""

from typing import Generator
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture(autouse=False)
def mock_sys_exit() -> Generator[MagicMock, None, None]:
    """Mock sys.exit to prevent it from actually exiting during tests"""
    with patch("sys.exit") as mock_exit:
        mock_exit.side_effect = SystemExit
        yield mock_exit
